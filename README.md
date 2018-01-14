# Microsoft SharePoint Online / Project Server Examples for Node.js

## Clone

```bash
git clone https://github.com/nsuvorov83/ms-sharepoint-project-nodejs.git && cd ms-sharepoint-project-nodejs
```

## Install dependencies

```bash
yarn install
```

or use `npm` instead.

## Configure connection with SharePoint

```bash
npm run config
```

Provide [connection parameters](https://github.com/s-KaiNet/node-sp-auth/wiki).

All the parameters are requested in interactive wizard-like approach and stored in `./config/private.json`.

It's a best practice to exclude files with creadentials from source control.

## Execute an example script

```bash
ts-node ./examples/[script-name]
```

## Bundling scripts

When creating automation scripts for production environment, e.g. Azure Job or Function or embedded application like Electron, it can be important to bundle and minify sources with positive performant effect as a result.

```bash
npm run pack
```

Initiates `webpack` bundling with files output to `./dist` folder. Only the content of the folder should be deployed to production environment.

`Dist` contains `index.js` with all dependencies bundled together brining huge performance improvements under cloud serverless scenarios.

JSOM scripts are located under `./dist/jsom` folder.

### Bundled script execution

```bash
node ./dist/index
```
