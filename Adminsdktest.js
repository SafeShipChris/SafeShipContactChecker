function testAdminDirectoryAccess() {
  const email = Session.getActiveUser().getEmail();
  const groups = AdminDirectory.Groups.list({ userKey: email }).groups || [];
  Logger.log(`${email} belongs to:`);
  groups.forEach(g => Logger.log('• ' + g.email));
}