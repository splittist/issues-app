import { describe, it, expect } from 'vitest'
import { render, screen } from '@testing-library/react'
import App from '../App'
import { CustomThemeProvider } from '../ThemeContext'

const renderApp = () => {
  return render(
    <CustomThemeProvider>
      <App />
    </CustomThemeProvider>
  )
}

describe('App', () => {
  it('should render the application title', () => {
    renderApp()
    expect(screen.getByText('Instant Issues List')).toBeInTheDocument()
  })

  it('should render the copyright footer', () => {
    renderApp()
    expect(screen.getByText('© 2025 | All liability comprehensively disclaimed')).toBeInTheDocument()
  })

  it('should render the main app structure', () => {
    renderApp()
    const app = screen.getByRole('banner') // header element
    expect(app).toBeInTheDocument()
  })
})