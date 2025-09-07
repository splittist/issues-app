import React, { createContext, useState, useEffect, ReactNode } from 'react';
import { ThemeProvider } from '@fluentui/react';
import { lightTheme, darkTheme } from './themes';

type ThemeMode = 'light' | 'dark';

interface ThemeContextType {
  themeMode: ThemeMode;
  toggleTheme: () => void;
}

export const ThemeContext = createContext<ThemeContextType | undefined>(undefined);

interface CustomThemeProviderProps {
  children: ReactNode;
}

export const CustomThemeProvider: React.FC<CustomThemeProviderProps> = ({ children }) => {
  const [themeMode, setThemeMode] = useState<ThemeMode>(() => {
    // Check for saved theme preference or default to system preference
    const savedTheme = localStorage.getItem('theme') as ThemeMode;
    if (savedTheme) {
      return savedTheme;
    }
    // Safely check for system preference with fallback
    try {
      return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
    } catch {
      return 'light'; // Default fallback
    }
  });

  useEffect(() => {
    localStorage.setItem('theme', themeMode);
    // Update CSS custom properties for non-FluentUI components
    document.documentElement.setAttribute('data-theme', themeMode);
  }, [themeMode]);

  const toggleTheme = () => {
    setThemeMode(prev => prev === 'light' ? 'dark' : 'light');
  };

  const currentTheme = themeMode === 'dark' ? darkTheme : lightTheme;

  return (
    <ThemeContext.Provider value={{ themeMode, toggleTheme }}>
      <ThemeProvider theme={currentTheme}>
        {children}
      </ThemeProvider>
    </ThemeContext.Provider>
  );
};