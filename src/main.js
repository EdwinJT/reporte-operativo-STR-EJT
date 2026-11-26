import { AppController } from './App.js';

const root = document.getElementById('root');
if (root) {
  const app = new AppController(root);
  app.start();
}
