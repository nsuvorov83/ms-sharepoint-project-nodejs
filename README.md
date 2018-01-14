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