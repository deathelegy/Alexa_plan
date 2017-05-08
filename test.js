const getmyContacts = () => new Promise((rs, rj) => {
    client.api('/me/contacts').then((contactsResult)=>{
      // get email address
      rs(contactsResult);
    }).catch((e) => {
      rj(e);
    })
  });


const sendEmail = (contactsResult) => new Promise((rs, rj) => {
  // handle contactsResult
  client.api('/me/sendMail').then((mailResult)=>{
    rs(mailResult);
  }).catch((e) => {
    rj(e);
  })
});

getmyContacts()
  .then(sendEmail)
  .then((mailResult)=>{
    // do something
    callback();
  }).catch((e) => {
    console.error('error happen', e);
  })
