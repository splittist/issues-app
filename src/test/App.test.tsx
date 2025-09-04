import { describe, it, expect } from 'vitest'
import { render, screen } from '@testing-library/react'
import App from '../App'

describe('App', () => {
  it('should render the application title', () => {
    render(<App />)
    expect(screen.getByText('Instant Issues List')).toBeInTheDocument()
  })

  it('should render the copyright footer', () => {
    render(<App />)
    expect(screen.getByText('© 2025 | All liability comprehensively disclaimed')).toBeInTheDocument()
  })

  it('should render the main app structure', () => {
    render(<App />)
    const app = screen.getByRole('banner') // header element
    expect(app).toBeInTheDocument()
  })
})