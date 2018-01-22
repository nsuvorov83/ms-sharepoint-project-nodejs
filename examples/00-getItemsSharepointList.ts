import * as path from 'path';
import { JsomNode, IJsomNodeInitSettings } from 'sp-jsom-node';

const settings = require(path.join(__dirname, '../config/private.json'));

let jsomNodeOptions: IJsomNodeInitSettings = {
  siteUrl: settings.siteUrlSP,
  authOptions: { ...settings }
};

interface IList {
  jur: string;
}

(async () => {

  (new JsomNode(jsomNodeOptions)).init();

  const ctx = SP.ClientContext.get_current();

  const oList = ctx.get_web().get_lists().getByTitle('Suppliers');
  const oView = oList.get_defaultView();
  const lookupListFields = oList.get_fields();

  let list = oList.getItems(new SP.CamlQuery()); //Делаем без фильтра, поэтому без set_viewXml

  ctx.load(list);
  ctx.load(lookupListFields);
  ctx.load(oView);

  await ctx.executeQueryPromise();

  let list_result: IList[] = list.get_data().map(p => {
    //Кодификация полей в SharePoint
    //Получение списка полей SP: https://your_domain.sharepoint.com/sites/your_site/_api/web/lists/GetByTitle('your_list_name')/fields
    return {
      jur: p.get_item("_x0043_ol1")
    };
  });
  
  console.log(list_result);

})().catch(console.log);
