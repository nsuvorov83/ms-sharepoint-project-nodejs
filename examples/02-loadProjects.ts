import * as path from 'path';
import { JsomNode, IJsomNodeInitSettings } from 'sp-jsom-node';

const settings = require(path.join(__dirname, '../config/private.json'));
declare const PS: any;

let jsomNodeOptions: IJsomNodeInitSettings = {
  siteUrl: settings.siteUrl,
  authOptions: { ...settings },
  modules: [ 'project' ]
};

interface IProject {
  Id: number;
  Name: string;
}

(async () => {

  (new JsomNode(jsomNodeOptions)).init();

  const projContext = PS.ProjectContext.get_current();
  const projectsCollection = projContext.get_projects();

  projContext.load(projectsCollection, 'Include(Id,Name)');

  await projContext.executeQueryPromise();

  let projects: IProject[] = projectsCollection.get_data().map(p => {
    return {
      Id: p.get_id(),
      Name: p.get_name()
    };
  });

  console.log('=== Projects: ===');
  console.log(projects);

})().catch(console.log);
