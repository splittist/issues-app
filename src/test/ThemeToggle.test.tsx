import { describe, it, expect, beforeEach } from 'vitest'
import { render, screen, fireEvent } from '@testing-library/react'
import { CustomThemeProvider } from '../ThemeContext'
import ThemeToggle from '../ThemeToggle'

const TestComponent = () => (
  <CustomThemeProvider>
    <ThemeToggle />
    <div data-testid="test-content">Test content</div>
  </CustomThemeProvider>
)

describe('ThemeToggle', () => {
  beforeEach(() => {
    // Clear localStorage before each test
    localStorage.clear()
  })

  it('should render the theme toggle component', () => {
    render(<TestComponent />)
    expect(screen.getByRole('switch', { name: 'Dark mode' })).toBeInTheDocument()
  })

  it('should toggle between light and dark themes', () => {
    render(<TestComponent />)
    const toggle = screen.getByRole('switch', { name: 'Dark mode' })
    
    // Should start in light mode
    expect(toggle).not.toBeChecked()
    expect(screen.getByText('Light')).toBeInTheDocument()
    
    // Click to switch to dark mode
    fireEvent.click(toggle)
    expect(toggle).toBeChecked()
    expect(screen.getByText('Dark')).toBeInTheDocument()
    
    // Click again to switch back to light mode
    fireEvent.click(toggle)
    expect(toggle).not.toBeChecked()
    expect(screen.getByText('Light')).toBeInTheDocument()
  })

  it('should update data-theme attribute on document element', () => {
    render(<TestComponent />)
    const toggle = screen.getByRole('switch', { name: 'Dark mode' })
    
    // Should start with light theme
    expect(document.documentElement.getAttribute('data-theme')).toBe('light')
    
    // Switch to dark mode
    fireEvent.click(toggle)
    expect(document.documentElement.getAttribute('data-theme')).toBe('dark')
    
    // Switch back to light mode
    fireEvent.click(toggle)
    expect(document.documentElement.getAttribute('data-theme')).toBe('light')
  })
})