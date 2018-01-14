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
  const lookupTables = projContext.get_lookupTables();

  const myLut = lookupTables.getByGuid('ebb134c6-29be-e711-80d3-00155da4760e');

  const myEntries = myLut.get_entries();

  const newId = SP.Guid.newGuid();

  const lutEntry = new PS.LookupEntryCreationInformation();

  lutEntry.set_description('my descr');
  lutEntry.set_sortIndex(10);
  lutEntry.set_id(newId);
  lutEntry.set_parentId(null);

  const lutValue = new PS.LookupEntryValue();
  lutValue.set_textValue('test');
  lutEntry.set_value(lutValue);
  myEntries.add(lutEntry);

  lookupTables.update(myEntries);

  projContext.load(lookupTables);
  projContext.load(myEntries);

  await projContext.executeQueryPromise();

  console.log('Done');

})().catch(console.log);
