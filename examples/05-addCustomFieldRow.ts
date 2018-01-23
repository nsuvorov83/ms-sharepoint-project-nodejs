import * as path from 'path';
import { JsomNode, IJsomNodeInitSettings } from 'sp-jsom-node';

const settings = require(path.join(__dirname, '../config/private.json'));

declare const PS: any;

let jsomNodeOptions: IJsomNodeInitSettings = {
  siteUrl: settings.msProjectUrl,
  authOptions: { ...settings },
  modules: [ 'project' ]
};

(async () => {

  (new JsomNode(jsomNodeOptions)).init();

  const projContext = PS.ProjectContext.get_current();
  const lookupTables = projContext.get_lookupTables();

  const myLut = lookupTables.getByGuid('ebb134c6-29be-e711-80d3-00155da4760e');

  const myEntries = myLut.get_entries();
  projContext.load(myEntries);

  //Отрабатываем промежуточные запросы
  await projContext.executeQueryPromise();

  //Определяем индекс новой записи и её GUID
  let new_index = myEntries.get_count() + 1;

  const newId = SP.Guid.newGuid();

  const lutEntry = new PS.LookupEntryCreationInformation();

  lutEntry.set_description('my descr22');
  lutEntry.set_sortIndex(new_index); //Index of new row (previosly get it see ex. 04-getCustomRows.ts)
  lutEntry.set_id(newId);
  lutEntry.set_parentId(null);

  const lutValue = new PS.LookupEntryValue();
  lutValue.set_textValue('test22');
  lutEntry.set_value(lutValue);
  myEntries.add(lutEntry);

  lookupTables.update(myEntries);

  await projContext.executeQueryPromise();

  console.log('Done');

})().catch(console.log);
