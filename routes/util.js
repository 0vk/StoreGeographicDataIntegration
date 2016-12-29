var ldap = require('ldapjs');
var client = ldap.createClient({
  url: 'ldap://172.18.1.70:389',
  timeout: 50,
  connectTimeout: 100
});

exports.authenticate = function(username, password) {
  var opts = {
    // filter: '(objectclass=*)',
    filter: '(sn=' + username + ')',
    scope: 'one',
    // attributes: '*'
    // This attribute list is what broke your solution
    // attributes: ['objectGUID','sAMAccountName','cn','mail','manager','memberOf']
  };
  return new Promise(function(resolve, reject) {
    try {
      client.bind('cn=' + username + ',ou=People,dc=sbuxcf,dc=com',
        password,
        function(
          error) {
          if (error) {
            // console.log(error.message);
            // client.unbind(function(error) {
            //   if (error) {
            //     console.log(error.message);
            //   } else {
            //     console.log('client disconnected');
            //   }
            // });
            reject(false);
          } else {
            resolve(true);
            // client.search('ou=People, dc=sbuxcf, dc=com', opts,
            //   function(
            //     error, search) {
            //     search.on('searchEntry', function(entry) {
            //       if (entry.object) {
            //         console.log('entry: ' + JSON.stringify(entry.object));
            //       }
            //       client.unbind(function(error) {
            //         if (error) {
            //           console.log(error.message);
            //         } else {
            //           console.log('client disconnected');
            //         }
            //       });
            //     });
            //
            //     search.on('error', function(error) {
            //       client.unbind(function(error) {
            //         if (error) {
            //           console.log(error.message);
            //         } else {
            //           console.log('client disconnected');
            //         }
            //       });
            //     });
            //
            //     // don't do this here
            //     //client.unbind(function(error) {if(error){console.log(error.message);} else{console.log('client disconnected');}});
            //   });
          }
        });
    } catch (error) {
      // console.log(error);
      // client.unbind(function(error) {
      //   if (error) {
      //     console.log(error.message);
      //   } else {
      //     console.log('client disconnected');
      //   }
      // });
      reject(false);
    }
  });
}
