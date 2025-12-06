
function isAuthorizedUser() {
  var activeUser = Session.getActiveUser();
  if (!activeUser) return false;

  var email = activeUser.getEmail();
  if (!email) return false;

  email = email.trim().toLowerCase();

  var authorizedUsers = [
    "admin@midlothianfriends.org",
    "clerk@midlothianfriends.org",
    "joemaddrey@gmail.com",
    "marcellavluna@gmail.com"
  ].map(function(u) { return u.toLowerCase(); });

  return authorizedUsers.indexOf(email) !== -1;
}

