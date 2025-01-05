import React from 'react';
import './App.css';
import WordHandler from './WordHandler';
import './WordHandler.css';

const App: React.FC = () => {
  return (
    <div className="App">
      <header className="App-header">
        <h1>Instant Issues List</h1>
        <WordHandler />
      </header>
    </div>
  );
};

export default App;
