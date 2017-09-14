# Build an Add-in with React

An add-in includes two parts, the web app and a manifest file.

### Step 1. Generate the React project by **Create React App**

Open Visual Studio Code, Click `View` -> `Integrated Terminal`.

In your terminal, input below code and press enter to go to Desktop folder.

```bash
cd Desktop
```

Generate your React app by

```bash
create-react-app my-addin
```


### Step 2. Generate the manifest file by **[Yo Office](https://github.com/OfficeDev/generator-office)**

Go to your app folder.

```bash
cd my-office-addin
```

Use the following command to create the Office manifest file with [the Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office):

```bash
yo office --skip-install
```

When prompted, supply the following information:

|Prompt|Response|
|---|---|
|New subfolder|No is default. Press enter or type 'n' to use current directory|
|Add-in name|Let's use the default name, just press enter|
|Supported Office host|Excel|
|Create new add-in|No, I only need a manifest file|

![Generate](./img/office-toolbox-generate.png)

> If prompted to overwrite package.json, type 'n' to decline.

### Step 3. Add and initialize Office.js

Type the following command into the terminal.

`code .` 

This will open your project in Visual Studio Code. Open the manifest and replace all `https://localhost:3000` to `http://localhost:3000`. The manifest filename ends with **manifest.xml** and is located in the root directory of your project.

Next, open **public/index.html**, and add the following before the `</head>` tag.

```html
<script src="https://appsforoffice.microsoft.com/lib/beta/hosted/office.debug.js"></script>
```

Open **src/index.js**, and replace `ReactDOM.render(<App />, document.getElementById('root'));` with the following:

```javascript
const Office = window.Office;

Office.initialize = () => {
  ReactDOM.render(<App />, document.getElementById('root'));
};
```

### Step 4. Add "Color Me" component

Open **src/App.js**. Replace contents with:

```javascript
import React, { Component } from 'react';

// const Excel = window.Excel;

class App extends Component {
  constructor(props) {
    super(props);

    this.onColorMe = this.onColorMe.bind(this);
  }

  onColorMe() {
    window.Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = 'green';
      await context.sync();
    });
  }

  render() {
    return (
      <button onClick={this.onColorMe}>Color Me</button>
    );
  }
}

export default App;
```

### Step 5. Run the app

Save all changes in VS Code. Reopen the terminal. Make sure you are in the root directory of the project, then run the dev server:

```bash
npm start
```

> If prompted, give Node.js permission to start the server. Otherwise, you won't be able to host your application.


### Step 6. Side load the manifest file into Office

To run the add-in, you need load the add-in into Excel. Below, we are using an open-source project currently in development, called [Office-Toolbox](https://github.com/OfficeDev/office-toolbox). It is not part of the official Office toolchain yet, but makes the sideloading process easier. Try office-toolbox below! Or you can follow our manual sideloading process documented [here](https://dev.office.com/docs/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).

#### Run Office-Toolbox

Open a new terminal window, and navigate to the root directory of the project (c:\Users\Administrator\Desktop\My-Office-Addin). Run the following command.

```bash
office-toolbox sideload -m my-office-add-in-manifest.xml -a excel
```

> **Don't Panic!** Office-Toolbox may spit out some errors, but it will still load your add-in into Excel.

Office-Toolbox will then launch Excel with your add-in loaded. Click the 'Show Taskpane' button on the 'Home' tab to reveal the taskpane.

![Final Result](img/final-colorme.png)

#### Congratulations! You just finished your first React add-in for Excel! 

> **Did You Know:** You can also run 'office-toolbox' without passing in arguments, and you will be prompted as shown in the image below.
![Sideload](./img/office-toolbox-sideload.png)
