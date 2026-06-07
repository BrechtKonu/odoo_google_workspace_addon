/**
 * EmailProcessor.gs
 * Handles Drive uploads for inline images and attachments before Odoo task/ticket creation.
 * All functions are best-effort — per-item errors are logged but never break task creation.
 */

/**
 * Main orchestrator. Call with a GmailMessage before the Odoo API call.
 * If no Drive folder is configured, falls back to cleanEmailHtml_ behavior.
 *
 * HTTP requests are batched into two UrlFetchApp.fetchAll calls to minimise latency:
 *   Phase 1 — MIME structure (if CID images present) + all Gmail URL image fetches, in parallel
 *   Phase 2 — CID attachment-ID fetches (only for parts without inline base64 data), in parallel
 *   Phase 3 — sequential Drive uploads (API does not support batch create)
 *
 * @param {GmailMessage} msg
 * @param {{ includeImages?: boolean, attachmentIndices?: number[]|null }} [options]
 *   includeImages   — When false, inline images are stripped but not uploaded. Default: true.
 *   attachmentIndices — Array of 0-based attachment indices to upload. null = all. Default: null.
 * @returns {{ emailBody: string, fileIds: string[], originalNames: string[] }}
 */
function processEmail(msg, options) {
  var opts = options || {};
  // Default: include all images; null attachmentIndices means "upload all"
  var includeImages = opts.includeImages !== false;
  var attachmentFilter = Array.isArray(opts.attachmentIndices) ? opts.attachmentIndices : null;

  var folderId = getDriveFolderId_();

  if (!folderId) {
    return {
      emailBody: cleanEmailHtml_(msg.getBody()),
      fileIds: [],
      originalNames: []
    };
  }

  // Resolve Drive folder once — getFolderById is a network round-trip
  var folder = DriveApp.getFolderById(folderId);
  var token = ScriptApp.getOAuthToken();
  var html = msg.getBody() || '';
  var msgId = msg.getId();

  // ── Collect work items ──────────────────────────────────────────────────

  var cidRefs = [];
  var gmailImgItems = [];

  if (includeImages) {
    var cidPattern = /src="cid:([^"]+)"/gi;
    var m;
    while ((m = cidPattern.exec(html)) !== null) {
      var cid = m[1].split('@')[0];
      if (cidRefs.indexOf(cid) === -1) cidRefs.push(cid);
    }

    var gmailImgPattern = /<img([^>]*)\ssrc="(https:\/\/mail\.google\.com\/[^"]+)"([^>]*)>/gi;
    while ((m = gmailImgPattern.exec(html)) !== null) {
      gmailImgItems.push({ before: m[1], src: m[2], after: m[3] });
    }
  }

  // ── Phase 1: fetchAll — MIME structure + Gmail URL images ───────────────

  var phase1Requests = [];
  var mimeIdx = -1;

  if (cidRefs.length > 0) {
    mimeIdx = phase1Requests.length;
    phase1Requests.push({
      url: 'https://gmail.googleapis.com/gmail/v1/users/me/messages/' + msgId + '?format=full',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
  }

  var gmailImgStartIdx = phase1Requests.length;
  gmailImgItems.forEach(function(item) {
    phase1Requests.push({
      url: item.src,
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
  });

  var phase1Responses = phase1Requests.length > 0
    ? UrlFetchApp.fetchAll(phase1Requests)
    : [];

  // ── Parse MIME → cidMap ─────────────────────────────────────────────────

  var cidMap = {};
  if (mimeIdx >= 0) {
    try {
      var mimeResp = phase1Responses[mimeIdx];
      if (mimeResp.getResponseCode() === 200) {
        cidMap = buildCidMap_(JSON.parse(mimeResp.getContentText()).payload);
      } else {
        console.error('processEmail: MIME fetch returned', mimeResp.getResponseCode());
      }
    } catch (e) {
      console.error('processEmail: error parsing MIME response', e);
    }
  }

  // ── Phase 2: fetchAll — CID attachment-ID fetches ───────────────────────
  // Parts with body.data are already available inline — only parts with
  // body.attachmentId need a second HTTP request.

  var cidInlineBlobs = {};   // cid → Blob  (decoded from inline base64 data)
  var cidAttachItems = [];   // [{ cid, mimeType, url }] needing HTTP fetch

  cidRefs.forEach(function(cid) {
    var part = cidMap[cid];
    if (!part) {
      console.error('processEmail: no MIME part for cid', cid);
      return;
    }
    var mimeType = part.mimeType || 'image/png';
    if (part.body && part.body.data) {
      try {
        var base64 = part.body.data.replace(/-/g, '+').replace(/_/g, '/');
        cidInlineBlobs[cid] = Utilities.newBlob(Utilities.base64Decode(base64), mimeType);
      } catch (e) {
        console.error('processEmail: inline CID decode error', cid, e);
      }
    } else if (part.body && part.body.attachmentId) {
      cidAttachItems.push({
        cid: cid,
        mimeType: mimeType,
        url: 'https://gmail.googleapis.com/gmail/v1/users/me/messages/' +
          msgId + '/attachments/' + part.body.attachmentId
      });
    }
  });

  var cidAttachBlobs = {};  // cid → Blob  (fetched via attachment-ID API)
  if (cidAttachItems.length > 0) {
    var phase2Responses = UrlFetchApp.fetchAll(cidAttachItems.map(function(item) {
      return { url: item.url, headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true };
    }));
    cidAttachItems.forEach(function(item, i) {
      try {
        var resp = phase2Responses[i];
        if (resp.getResponseCode() !== 200) {
          console.error('processEmail: CID attachment fetch non-200', item.cid, resp.getResponseCode());
          return;
        }
        var data = JSON.parse(resp.getContentText()).data;
        cidAttachBlobs[item.cid] = Utilities.newBlob(
          Utilities.base64Decode(data.replace(/-/g, '+').replace(/_/g, '/')),
          item.mimeType
        );
      } catch (e) {
        console.error('processEmail: CID attachment decode error', item.cid, e);
      }
    });
  }

  // ── Odoo-attachment mode ────────────────────────────────────────────────
  // Instead of uploading to Drive, return the fetched bytes (base64) so the
  // Odoo controller can rehost them as ir.attachment on the record. Inline
  // images keep a cid sentinel in the body for the controller to rewrite.
  if (opts.mode === 'odoo') {
    return collectOdooAttachments_(
      html, cidRefs, cidInlineBlobs, cidAttachBlobs,
      gmailImgItems, phase1Responses, gmailImgStartIdx, msg, attachmentFilter
    );
  }

  // ── Phase 3: Drive uploads ──────────────────────────────────────────────

  var fileIds = [];
  var originalNames = [];
  var cidToUrl = {};      // cid → Drive URL (for HTML rewriting)
  var gmailSrcToUrl = {}; // original Gmail src → Drive URL (for HTML rewriting)
  var now = Date.now();

  // CID images
  cidRefs.forEach(function(cid, idx) {
    var blob = cidInlineBlobs[cid] || cidAttachBlobs[cid];
    if (!blob) return;
    try {
      var contentType = blob.getContentType() || 'image/png';
      var ext = contentType.split('/')[1] || 'png';
      var originalName = 'cid_image_' + idx + '.' + ext;
      var uploaded = uploadFileToDrive_(
        blob, 'tmp_' + now + '_cid_' + idx + '_' + originalName, true, folder
      );
      fileIds.push(uploaded.fileId);
      originalNames.push(originalName);
      cidToUrl[cid] = uploaded.driveUrl;
    } catch (e) {
      console.error('processEmail: Drive upload error for CID', cid, e);
    }
  });

  // Gmail URL images
  gmailImgItems.forEach(function(item, idx) {
    try {
      var resp = phase1Responses[gmailImgStartIdx + idx];
      if (resp.getResponseCode() !== 200) {
        console.error('processEmail: Gmail image fetch non-200', idx, resp.getResponseCode());
        return;
      }
      var blob = resp.getBlob();
      var contentType = blob.getContentType() || 'image/png';
      var ext = contentType.split('/')[1] || 'png';
      var originalName = 'inline_image_' + idx + '.' + ext;
      var uploaded = uploadFileToDrive_(
        blob, 'tmp_' + now + '_gmailimg_' + idx + '_' + originalName, true, folder
      );
      fileIds.push(uploaded.fileId);
      originalNames.push(originalName);
      gmailSrcToUrl[item.src] = uploaded.driveUrl;
    } catch (e) {
      console.error('processEmail: Drive upload error for Gmail image', idx, e);
    }
  });

  // Email attachments (filtered by selection if attachmentFilter is set)
  try {
    msg.getAttachments().forEach(function(attachment, idx) {
      if (attachmentFilter !== null && attachmentFilter.indexOf(idx) === -1) return;
      try {
        var name = attachment.getName() || ('attachment_' + idx);
        var uploaded = uploadFileToDrive_(
          attachment.copyBlob(), 'tmp_' + now + '_att_' + idx + '_' + name, false, folder
        );
        fileIds.push(uploaded.fileId);
        originalNames.push(name);
      } catch (e) {
        console.error('processEmail: attachment upload error', idx, e);
      }
    });
  } catch (e) {
    console.error('processEmail: getAttachments failed', e);
  }

  // ── Phase 4: rewrite HTML ───────────────────────────────────────────────

  var resultHtml = html
    .replace(/src="cid:([^"]+)"/gi, function(match, rawCid) {
      var cid = rawCid.split('@')[0];
      return cidToUrl[cid] ? 'src="' + cidToUrl[cid] + '"' : match;
    })
    .replace(/<img([^>]*)\ssrc="(https:\/\/mail\.google\.com\/[^"]+)"([^>]*)>/gi,
      function(match, before, src, after) {
        return gmailSrcToUrl[src]
          ? '<img' + before + ' src="' + gmailSrcToUrl[src] + '"' + after + '>'
          : match;
      }
    );

  return {
    emailBody: cleanEmailHtml_(resultHtml),
    fileIds: fileIds,
    originalNames: originalNames
  };
}

/**
 * Builds the Odoo-attachment payload for processEmail's 'odoo' mode.
 * Returns { emailBody, attachments } where attachments is a list of
 * { name, mimetype, data (base64), cid } — cid set for inline images
 * (matched against a src="cid:konu-img-N" sentinel in the body), null for
 * file attachments. No Drive upload happens in this mode.
 */
function collectOdooAttachments_(html, cidRefs, cidInlineBlobs, cidAttachBlobs,
                                 gmailImgItems, phase1Responses, gmailImgStartIdx, msg, attachmentFilter) {
  var attachments = [];
  var body = html;
  var imgIndex = 0;

  function pushImage_(blob, konuCid) {
    var ct = blob.getContentType() || 'image/png';
    var ext = (ct.split('/')[1] || 'png').split(';')[0];
    attachments.push({
      name: 'image_' + (imgIndex + 1) + '.' + ext,
      mimetype: ct,
      data: Utilities.base64Encode(blob.getBytes()),
      cid: konuCid
    });
  }

  // Inline CID images — rewrite src="cid:<raw>" (raw may carry @domain) to a
  // stable sentinel the controller recognises.
  cidRefs.forEach(function(cid) {
    var blob = cidInlineBlobs[cid] || cidAttachBlobs[cid];
    if (!blob) return;
    var konuCid = 'konu-img-' + imgIndex;
    try {
      pushImage_(blob, konuCid);
      var re = new RegExp('src=["\']cid:' + escapeRegExp_(cid) + '[^"\']*["\']', 'gi');
      body = body.replace(re, 'src="cid:' + konuCid + '"');
      imgIndex++;
    } catch (e) {
      console.error('collectOdooAttachments_: CID image error', cid, e);
    }
  });

  // Inline Gmail-URL images — fetched in phase 1.
  gmailImgItems.forEach(function(item, idx) {
    var resp = phase1Responses[gmailImgStartIdx + idx];
    if (!resp || resp.getResponseCode() !== 200) return;
    var konuCid = 'konu-img-' + imgIndex;
    try {
      pushImage_(resp.getBlob(), konuCid);
      body = body.split('src="' + item.src + '"').join('src="cid:' + konuCid + '"');
      imgIndex++;
    } catch (e) {
      console.error('collectOdooAttachments_: Gmail image error', idx, e);
    }
  });

  // File attachments (filtered by selection if attachmentFilter is set).
  try {
    msg.getAttachments().forEach(function(att, idx) {
      if (attachmentFilter !== null && attachmentFilter.indexOf(idx) === -1) return;
      try {
        var blob = att.copyBlob();
        attachments.push({
          name: att.getName() || ('attachment_' + idx),
          mimetype: blob.getContentType() || '',
          data: Utilities.base64Encode(blob.getBytes()),
          cid: null
        });
      } catch (e) {
        console.error('collectOdooAttachments_: attachment error', idx, e);
      }
    });
  } catch (e) {
    console.error('collectOdooAttachments_: getAttachments failed', e);
  }

  return { emailBody: cleanEmailHtml_(body, true), attachments: attachments };
}

/**
 * Recursively walks a Gmail MIME payload tree and returns a map of
 * Content-ID → MIME part object.
 *
 * @param {Object} payload  Gmail MIME payload (or part)
 * @returns {Object}  { cid: mimePartObject, ... }
 */
function buildCidMap_(payload) {
  var map = {};
  if (!payload) return map;

  var headers = payload.headers || [];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].name.toLowerCase() === 'content-id') {
      var rawCid = headers[i].value.trim().replace(/^<|>$/g, '');
      var key = rawCid.split('@')[0];
      if (key) map[key] = payload;
    }
  }

  var parts = payload.parts || [];
  for (var j = 0; j < parts.length; j++) {
    var subMap = buildCidMap_(parts[j]);
    for (var k in subMap) {
      if (!map[k]) map[k] = subMap[k];
    }
  }

  return map;
}

/**
 * Uploads a blob to the configured Drive folder.
 * Inline images are shared publicly and served via Google's image CDN so that
 * Odoo chatter can render them without authentication redirects.
 * Attachments keep domain-only sharing with the standard Drive viewer URL.
 *
 * @param {Blob} blob
 * @param {string} filename        Temporary filename to use on Drive
 * @param {boolean} isInlineImage  When true, use ANYONE_WITH_LINK + lh3 CDN URL
 * @param {DriveFolder} folder     Pre-resolved Drive folder
 * @returns {{ fileId: string, driveUrl: string }}
 */
function uploadFileToDrive_(blob, filename, isInlineImage, folder) {
  blob.setName(filename);
  var file = folder.createFile(blob);
  var fileId = file.getId();

  if (isInlineImage) {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return {
      fileId: fileId,
      driveUrl: 'https://lh3.googleusercontent.com/d/' + fileId
    };
  } else {
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    return {
      fileId: fileId,
      driveUrl: 'https://drive.google.com/uc?export=view&id=' + fileId
    };
  }
}

/**
 * Renames and moves uploaded files into year/month subfolders after Odoo returns the record ID.
 * Best-effort — errors are logged but never thrown.
 *
 * @param {string[]} fileIds
 * @param {string[]} originalNames
 * @param {{ year: string, month: string, recordId: string, recordType: string }} context
 */
function renameAndMoveFiles_(fileIds, originalNames, context) {
  if (!fileIds || fileIds.length === 0) return;

  var folderId = getDriveFolderId_();
  if (!folderId) return;

  var rootFolder = DriveApp.getFolderById(folderId);
  var yearFolder = getOrCreateSubfolder_(rootFolder, context.year);
  var monthFolder = getOrCreateSubfolder_(yearFolder, context.month);

  for (var i = 0; i < fileIds.length; i++) {
    try {
      var file = DriveApp.getFileById(fileIds[i]);
      var originalName = originalNames[i] || ('file_' + i);
      file.setName(context.recordType + '_' + context.recordId + '_' + originalName);
      file.moveTo(monthFolder);
    } catch (e) {
      console.error('renameAndMoveFiles_: error on file', fileIds[i], e);
    }
  }
}

/**
 * Gets or creates a named subfolder inside a parent Drive folder.
 *
 * @param {DriveFolder} parent
 * @param {string} name
 * @returns {DriveFolder}
 */
function getOrCreateSubfolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
