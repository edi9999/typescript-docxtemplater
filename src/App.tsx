import React from 'react';
import logo from './logo.svg';
import './App.css';
import Docxtemplater from "docxtemplater";
import expressionParser from "docxtemplater/expressions";

import PizZip from 'pizzip';
import PizZipUtils from 'pizzip/utils/index.js';

import { saveAs } from 'file-saver';

function loadFile(url: string, callback: any) {
  PizZipUtils.getBinaryContent(url, callback);
}

function generate() {
  loadFile(
      'https://docxtemplater.com/tag-example.docx',
      function (error: Error | null, content: string) {
        if (error) {
          throw error;
        }
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
          paragraphLoop: true,
          linebreaks: true,
          parser: expressionParser,
        });
        doc.setData({
          first_name: 'John',
          last_name: 'Doe',
          phone: '0652455478',
          description: 'New Website',
        });
        try {
          // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
          doc.render();
        } catch (error: any) {
          // The error thrown here contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
          if (error.properties && error.properties.errors instanceof Array) {
            const errorMessages = error.properties.errors
              .map(function (error: any) {
                return error.properties.explanation;
              })
              .join('\n');
            console.log('errorMessages', errorMessages);
            // errorMessages is a humanly readable message looking like this :
            // 'The tag beginning with "foobar" is unopened'
          }
          throw error;
        }
        const out = doc.getZip().generate({
          type: 'blob',
          mimeType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });
        // Output the document using Data-URI
        saveAs(out, 'output.docx');
})
}

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <p>
          Edit <code>src/App.tsx</code> and save to reload.
        </p>
        <a
          className="App-link"
          href="https://reactjs.org"
          target="_blank"
          rel="noopener noreferrer"
        >
          Learn React
        </a>
        <button onClick={()=>generate()}>
        Foobar
        </button>
      </header>
    </div>
  );
}

export default App;
