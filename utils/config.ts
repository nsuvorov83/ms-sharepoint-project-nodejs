import { AuthConfig } from 'node-sp-auth-config';
import { JsomNode } from 'sp-jsom-node';

const authConfig = new AuthConfig({
  // configPath: './config/private.json', // Default is used over examples library
  encryptPassword: true,
  saveConfigOnDisk: true,
  forcePrompts: true
});

authConfig.getContext().then(async context => {
  (new JsomNode(context)).init();
  let ctx = SP.ClientContext.get_current();
  let oWeb = ctx.get_web();
  ctx.load(oWeb, 'Title');
  console.log(`\nTrying to connect to ${context.siteUrl}`);
  await ctx.executeQueryPromise();
  console.log(`Connected to web '${oWeb.get_title()}'\n`);
})
.catch(console.log);
