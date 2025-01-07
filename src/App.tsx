import React from 'react';
import { Text } from '@fluentui/react/lib/Text';
import './App.css';
import WordHandler from './WordHandler';
import './WordHandler.css';

const App: React.FC = () => {
  return (
    <div className="App">
      <header className="App-header">
        <Text variant='xxLarge'>Instant Issues List</Text>
        <WordHandler />
      </header>
      <footer className="App-footer">
        <Text variant='small'>© 2025 | All liability comprehensively disclaimed</Text>
      </footer>
    </div>
  );
};

export default App;
