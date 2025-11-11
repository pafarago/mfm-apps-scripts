function getWpConfig_() {
  const props = PropertiesService.getScriptProperties();
  const base = props.getProperty('WP_BASE_URL');
  const user = props.getProperty('WP_USER');
  const pass = props.getProperty('WP_APP_PASSWORD');
  const cat  = props.getProperty('WP_CATEGORY_ID');
  if (!base || !user || !pass) return null;
  return {
    baseUrl: base.replace(/\/+$/, ''),
    user,
    pass,
    categoryId: cat ? Number(cat) : null
  };
}
/***** Post News to midlothianfriends.org *****/

function postNewsletterToWordPress_(subject, htmlBody) {
  const wp = getWpConfig_();
  if (!wp) {
    SpreadsheetApp.getUi().alert('WP config missing (check Script properties).');
    return;
  }

  const url = wp.baseUrl + '/wp-json/wp/v2/posts';

  const payload = {
    title: subject,
    content: htmlBody,
    status: 'publish'
  };
  if (wp.categoryId) payload.categories = [wp.categoryId];

  const auth = Utilities.base64Encode(wp.user + ':' + wp.pass);

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Basic ' + auth },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('WP post failed: ' + code + ' ' + res.getContentText());
  }
}

function menuPostToNews() {
  const cfg = getConfig_();
  const {rows} = getRows_();

  // Use the first row with Send? = TRUE, or fall back to the first row
  const row = rows.find(r => asBool_(r['Send?'])) || rows[0];
  if (!row) {
    SpreadsheetApp.getUi().alert('No rows available in MailingList.');
    return;
  }

  const merged = renderDocToBlobs_(cfg.DOC_TEMPLATE_ID, row);
  
  // If template is personalized, consider making it generic:
  // Replace any personalized greeting with a generic one

  const htmlBody = merged.htmlBody.replace(
  /Dear[\s\n]+[A-Z][A-Za-z.'\s-]*[,:\n]?/i,
  'Dear Member of the Midlothian Friends Community,'
);
  try {
    postNewsletterToWordPress_(cfg.SUBJECT, htmlBody);
    SpreadsheetApp.getUi().alert('Newsletter posted to website.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Failed to post to website: ' + e.message);
  }
}
