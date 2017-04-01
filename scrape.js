if (process.argv.length < 6) {
  console.log('usage: node scrape username password tenant_id filename [next_link]');
  return;
}

const
  msRestAzure = require('ms-rest-azure'),
  graphRbacManagementClient = require('azure-graph'),
  fs = require('fs'),
  username = process.argv[2],
  password = process.argv[3],
  tenantId = process.argv[4],
  filename = process.argv[5],
  initLink = process.argv[6] || null;

const login = (username, password) =>
  new Promise((resolve, reject) =>
    msRestAzure.loginWithUsernamePassword(username, password, {tokenAudience: 'graph', domain: tenantId}, (err, credentials, subscriptions) => err ? reject(err) : resolve({credentials, subscriptions})));

const queryUsers = (client, nextLink = null) =>
  new Promise((resolve, reject) =>
    nextLink
      ? client.users.listNext(nextLink, (err, result, request, response) =>
          err ? reject(err) : resolve({result: [...result], nextLink: result.odatanextLink}))
      : client.users.list((err, result, request, response) =>
          err ? reject(err) : resolve({result: [...result], nextLink: result.odatanextLink})));

async function scrape() {
  const
    client = new graphRbacManagementClient((await login(username, password)).credentials, tenantId),
    out = fs.createWriteStream('./' + filename, {flags: 'w'});
  let
    result,
    nextLink = initLink;
  while (({result, nextLink} = await queryUsers(client, nextLink)) && nextLink) {
    result.forEach(e => out.write(e.objectId + '\t' + e.objectType + '\t' + e.userPrincipalName + '\t' + e.displayName + '\t' + e.mail + '\t' + e.mailNickname + '\n'));
    console.log(nextLink);
  }
  out.end();
}

async function main() {
  console.log('start: ' + new Date());
  try {
    await scrape();
  } catch (err) {
    console.error(err);
    console.log('pass the last link as a parameter to restart from this point');
  }
  console.log('end: ' + new Date());
};

main();
