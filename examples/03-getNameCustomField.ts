import * as path from 'path';
import { JsomNode, IJsomNodeInitSettings } from 'sp-jsom-node';

const settings = require(path.join(__dirname, '../config/private.json'));
declare const PS: any;

let jsomNodeOptions: IJsomNodeInitSettings = {
  siteUrl: settings.siteUrl,
  authOptions: { ...settings },
  modules: [ 'project' ]
};

(async () => {

  (new JsomNode(jsomNodeOptions)).init();

  const projContext = PS.ProjectContext.get_current();

  const customFields = projContext.get_customFields();
  const jurCustomField = customFields.getByGuid('09e501e7-29be-e711-80ea-00155da4370e');

  // const projectLookupTableEntries = jurCustomField.get_lookupEntries();

  projContext.load(jurCustomField);

  await projContext.executeQueryPromise();

  console.log('=== Custom field: ===');
  console.log(jurCustomField.get_name());

})().catch(console.log);
