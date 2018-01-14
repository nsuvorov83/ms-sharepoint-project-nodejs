import { JsomNode } from 'sp-jsom-node';

(new JsomNode()).wizard().then(async settings => {

  const ctx = SP.ClientContext.get_current();
  const oList = ctx.get_web().get_lists().getByTitle('New Lists');

  const itemCreateInfo = new SP.ListItemCreationInformation();
  const oListItem = oList.addItem(itemCreateInfo);

  oListItem.set_item('Title', 'my record');

  oListItem.update();
  ctx.load(oListItem);

  await ctx.executeQueryPromise();

  console.log(`Item has been created, ID ${oListItem.get_id()}`);

}).catch(console.log);
