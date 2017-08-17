# Build an Add-in with React

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

### Step 2. Generate the manifest file by **Office Toolbox**.

Go to your app folder.

```bash
cd my-addin
```

Generate the manifest file following the steps below.

```bash
office-toolbox
```

![Generate](./img/office-toolbox-generate.png)

You should be able to see your manifest file with the name ends with **manifest.xml**.

### Step 3. Prepare

Open **public/index.html**, add

```html
<script src="https://appsforoffice.microsoft.com/lib/beta/hosted/office.debug.js"></script>
```

before `</head>` tag.

Open **src/index.js**, add `Office.initialize` out of `ReactDOM.render(<App />, document.getElementById('root'));` like below:

```typescript
Office.initialize = () => {
  ReactDOM.render(<App />, document.getElementById('root'));
};
```

### Step 4. Add "Color Me"

Open **src/App.js**. Replace by

```javascript
import React, { Component } from 'react';

const Excel = window.Excel;

class App extends Component {
  constructor(props) {
    super(props);

    this.onColorMe = this.onColorMe.bind(this);
  }

  onColorMe() {
    Excel.run(async (context) => {
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

### Step 5. Run

Run the dev server through the terminal.

```bash
npm start
```

### Step 6. Side load

To run the add-in, you need side-load the add-in within the Excel application.

Run this in terminal and following the steps below.

```bash
office-toolbox
```

![Sideload](./img/office-toolbox-sideload.png)

Congratulations you just finish your first React add-in for Excel!

