import * as path from 'path';
import { JsomNode, IJsomNodeInitSettings } from 'sp-jsom-node';

const settings = require(path.join(__dirname, '../config/private.json'));

let jsomNodeOptions: IJsomNodeInitSettings = {
  siteUrl: settings.siteUrl,
  authOptions: { ...settings },
  modules: [ 'project' ]
};

(async () => {
  
    (new JsomNode(jsomNodeOptions)).init();

  const ctx = SP.ClientContext.get_current();
  const oList = ctx.get_web().get_lists().getByTitle('New Lists');

  const itemCreateInfo = new SP.ListItemCreationInformation();
  const oListItem = oList.addItem(itemCreateInfo);

  oListItem.set_item('Title', 'my record1343434');

  oListItem.update();
  ctx.load(oListItem);

  await ctx.executeQueryPromise();

  console.log(`Item has been created, ID ${oListItem.get_id()}`);

})().catch(console.log);  
