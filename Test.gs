// source: https://www.pbainbridge.co.uk/2021/04/extract-list-of-google-group-members.html
function myFunction() {
  const group = GroupsApp.getGroupByEmail('all@link-cc.co.jp');
  const members = group.getUsers();
  // console.log(members)
  
  let list = [];
  for(member of members){
    console.log(member.getEmail())
  }
}
