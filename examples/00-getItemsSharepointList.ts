import * as path from 'path';
import { JsomNode, IJsomNodeInitSettings } from 'sp-jsom-node';

const settings = require(path.join(__dirname, '../config/private.json'));

let jsomNodeOptions: IJsomNodeInitSettings = {
  siteUrl: settings.siteUrlSP,
  authOptions: { ...settings }
};

interface IList {
  jur: string;
  email: string;
}

interface IFields {
  name: string;
}

(async () => {

  (new JsomNode(jsomNodeOptions)).init();

  const ctx = SP.ClientContext.get_current();

  const oList = ctx.get_web().get_lists().getByTitle('Suppliers');
  const oView = oList.get_defaultView();
  const lookupListFields = oList.get_fields();


  var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(
        "<View>" +
        "</View>");

  let list = oList.getItems(camlQuery);

  ctx.load(list);
  ctx.load(lookupListFields);
  ctx.load(oView);

  await ctx.executeQueryPromise();

  let fields_result: IFields[] = lookupListFields.get_data().map(f => {
    return {
      name: f.get_title()
    }
  });

  let list_result: IList[] = list.get_data().map(p => {
    //Кодификация полей в SharePoint. Чтобы узнать эти коды нужно импортировать список в Excel
    return {
      jur: p.get_item("Jur.face"),
      email: p.get_item("Email")
    };
  });

  
  console.log(list_result);
  console.log(fields_result);

})().catch(console.log);
