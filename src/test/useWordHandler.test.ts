import type { FormEvent } from 'react'
import { act, renderHook } from '@testing-library/react'
import { describe, expect, it, vi, beforeEach } from 'vitest'
import { Packer } from 'docx'
import { saveAs } from 'file-saver'
import { useWordHandler } from '../useWordHandler'
import type { ExtractedParagraph } from '../types'

const extractParagraphsMock = vi.fn()
const buildSectionsMock = vi.fn()
const buildStylesMock = vi.fn()

vi.mock('../wordUtils', () => ({
  extractParagraphs: (...args: unknown[]) => extractParagraphsMock(...args),
  buildSections: (...args: unknown[]) => buildSectionsMock(...args),
  buildStyles: (...args: unknown[]) => buildStylesMock(...args),
}))

vi.mock('file-saver', () => ({
  saveAs: vi.fn(),
}))

describe('useWordHandler', () => {
  beforeEach(() => {
    extractParagraphsMock.mockReset()
    buildSectionsMock.mockReset()
    buildStylesMock.mockReset()
    vi.restoreAllMocks()
  })

  it('should add, move, and remove files', () => {
    const fileA = new File(['a'], 'a.docx')
    const fileB = new File(['b'], 'b.docx')
    const { result } = renderHook(() => useWordHandler())

    act(() => {
      result.current.addFiles([fileA, fileB])
    })

    expect(result.current.files.map(file => file.name)).toEqual(['a.docx', 'b.docx'])

    act(() => {
      result.current.moveFile(0, 1)
    })

    expect(result.current.files.map(file => file.name)).toEqual(['b.docx', 'a.docx'])

    act(() => {
      result.current.removeFile(0)
    })

    expect(result.current.files.map(file => file.name)).toEqual(['a.docx'])
  })

  it('should update criteria and output file name', () => {
    const { result } = renderHook(() => useWordHandler())

    act(() => {
      result.current.handleCriteriaChange({
        target: { name: 'comments' }
      } as unknown as FormEvent<HTMLInputElement>, true)
    })

    act(() => {
      result.current.handleOutputFileNameChange(undefined, 'custom-report.docx')
    })

    expect(result.current.criteria.comments).toBe(true)
    expect(result.current.outputFileName).toBe('custom-report.docx')
  })

  it('should run the save pipeline with extracted paragraphs and output file name', async () => {
    const file = new File(['content'], 'sample.docx')
    const extractedParagraphs = [[{ paragraph: {} as ExtractedParagraph['paragraph'], comments: [], source: 'document' }]]
    const mockSections = [{ properties: {}, children: [] }]
    const mockStyles = { paragraphStyles: [], characterStyles: [] }
    const blob = new Blob(['result'])

    extractParagraphsMock.mockResolvedValue(extractedParagraphs[0])
    buildSectionsMock.mockReturnValue(mockSections)
    buildStylesMock.mockReturnValue(mockStyles)
    vi.spyOn(Packer, 'toBlob').mockResolvedValue(blob)

    const { result } = renderHook(() => useWordHandler())

    act(() => {
      result.current.addFiles([file])
      result.current.handleOutputFileNameChange(undefined, 'issues-report.docx')
    })

    await act(async () => {
      await result.current.handleSaveFile()
    })

    expect(extractParagraphsMock).toHaveBeenCalledWith(file, result.current.criteria)
    expect(buildSectionsMock).toHaveBeenCalledWith(extractedParagraphs, ['sample.docx'])
    expect(buildStylesMock).toHaveBeenCalled()
    expect(Packer.toBlob).toHaveBeenCalled()
    expect(saveAs).toHaveBeenCalledWith(blob, 'issues-report.docx')
  })
})
