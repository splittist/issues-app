import React from 'react';
import { Toggle } from '@fluentui/react';
import { useTheme } from './useTheme';

const ThemeToggle: React.FC = () => {
  const { themeMode, toggleTheme } = useTheme();

  return (
    <Toggle
      label="Dark mode"
      checked={themeMode === 'dark'}
      onChange={toggleTheme}
      onText="Dark"
      offText="Light"
      styles={{
        root: {
          marginBottom: '10px',
        },
      }}
    />
  );
};

export default ThemeToggle;