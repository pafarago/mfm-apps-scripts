/***** MENU *****/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Newsletter')
    .addItem('Process subscription requests', 'menuProcessSubscriptions')
    .addItem('Refresh mailing list from roster', 'menuRefreshMailingList')
    .addItem('Validate template vs sheet', 'menuValidateTemplate')
    .addItem('Preview to me', 'menuPreview')
    .addSeparator()
    .addItem('Send to all (throttled)', 'menuSendAll')
    .addItem('Post latest newsletter to website', 'menuPostToNews')
    .addToUi();
}
