/*** Drive.gs — Google Drive 이미지 읽기 (고급 서비스 Drive) ***/

function DRIVE_extractId(urlOrId) {
  if (!urlOrId) return null;
  const s = String(urlOrId);

  let m = s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/); if (m) return m[1];
  m = s.match(/\/folders\/([a-zA-Z0-9_-]{10,})(?:[/?#]|$)/); if (m) return m[1];
  m = s.match(/\/file\/d\/([a-zA-Z0-9_-]{10,})(?:[/?#]|$)/);  if (m) return m[1];
  m = s.match(/[-\w]{25,}/); return m ? m[0] : null;
}

/** RichTextValue에서 첫 링크 안전 추출(부분 링크 대응) */
function getFirstLinkUrlFromRich_(rtv) {
  try {
    if (!rtv) return null;
    if (rtv.getRuns) {
      const runs = rtv.getRuns();
      for (const run of runs) {
        const u = run.getLinkUrl();
        if (u) return u;
      }
    }
    if (rtv.getLinkUrl) return rtv.getLinkUrl();
  } catch (_) {}
  return null;
}

/** ID가 폴더인지 확인 + 파일이면 첫 부모 폴더로 보정 (v2/v3 호환) */
function _ensureFolderId(id) {
  const res = Drive.Files.get(id, {
    // v3: parents = string[]
    // v2: parents = [{ id: ... }]
    fields: 'id,mimeType,parents',
    supportsAllDrives: true,   // v3
    supportsTeamDrives: true   // v2
  });

  if (!res) return null;

  // 폴더면 그대로 반환
  if (res.mimeType === 'application/vnd.google-apps.folder') return res.id;

  // 파일이면 첫 부모 폴더 ID 반환 (v2/v3 호환 처리)
  const parents = res.parents || [];
  if (!parents.length) return null;

  const first = parents[0];
  const parentId = (typeof first === 'string') ? first : (first.id || first);
  return parentId || null;
}

/**
 * 폴더(ID/URL) 안의 이미지 파일 나열 (페이징 + 공유드라이브 지원) — v3 전용
 * 반환: [{id,name,mimeType,url,thumb,viewUrl}]
 */
function listImagesInFolder(urlOrId) {
  const rawId = DRIVE_extractId(urlOrId);
  if (!rawId) throw new Error('유효하지 않은 폴더/파일 URL 또는 ID입니다: ' + urlOrId);

  // 파일ID일 수 있으므로 폴더ID로 보정
  const folderId = _ensureFolderId(rawId);
  if (!folderId) throw new Error('폴더를 찾을 수 없습니다(파일이거나 접근권한 없음): ' + rawId);

  const query = `'${folderId}' in parents and trashed = false and mimeType contains 'image/'`;
  const out = [];
  let pageToken = null;

  do {
    const params = {
      q: query,
      pageSize: 200,
      pageToken: pageToken,
      // 공유드라이브(Shared drives) 포함
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      // v3 필드만 지정
      fields: 'files(id,name,mimeType,webViewLink,thumbnailLink,imageMediaMetadata(width,height)),nextPageToken'
    };

    const res = Drive.Files.list(params);
    const files = (res && res.files) || [];
    pageToken = res && res.nextPageToken;

    for (const f of files) {
      const id    = f.id;
      const name  = f.name || '';
      const mime  = f.mimeType || '';
      const w     = (f.imageMediaMetadata && f.imageMediaMetadata.width)  || null;
      const h     = (f.imageMediaMetadata && f.imageMediaMetadata.height) || null;
      const view  = f.webViewLink || '';
      const thumb = 'https://drive.google.com/thumbnail?id=' + id;
      const url   = 'https://drive.google.com/uc?export=view&id=' + id;          // 미리보기

      out.push({
        id,
        name,
        mimeType: mime,
        url,           // 필요하면 써요
        thumb,         // ★ 이걸 갤러리 src로 사용
        viewUrl: 'https://drive.google.com/file/d/' + id + '/view',
        width: w, height: h
      });
    }
  } while (pageToken);

  return out;
}

/** (옵션) 시트 C열 링크로 바로 호출 */
function listImagesForListing(folderUrlInSheet) {
  if (!folderUrlInSheet) return [];
  return listImagesInFolder(folderUrlInSheet);
}

/** 테스트 */
function _testListImages() {
  const folderUrl = 'https://drive.google.com/drive/folders/1EYhttBPY254HBAQX_k3GCbqrVLiYLcIf';
  const list = listImagesInFolder(folderUrl);
  Logger.log(JSON.stringify(list, null, 2));
}
