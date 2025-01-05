import React, { useState, ChangeEvent, FormEvent } from 'react';
import { saveAs } from 'file-saver';
import { Document, 
  Packer, 
  } from 'docx';

import { DndProvider } from 'react-dnd';
import { HTML5Backend } from 'react-dnd-html5-backend';
import { initializeIcons } from '@fluentui/font-icons-mdl2';
import { Checkbox, Stack, TextField, PrimaryButton, } from '@fluentui/react';
import { useDropzone } from 'react-dropzone';

import FileItem from './FileItem';
import { ExtractedParagraph, Criteria } from './types';
import { buildSections, buildStyles, extractParagraphs } from './wordUtils';
import { dateToday } from './utils';

initializeIcons(); // For checkmark icon

/**
 * Main component for handling Word document processing.
 * @returns A React component.
 */
const WordHandler: React.FC = () => {
  const [files, setFiles] = useState<File[]>([]);
  const [criteria, setCriteria] = useState<Criteria>({ redline: true, highlight: false, squareBrackets: false, comments: false });
  const [outputFileName, setOutputFileName] = useState<string>('report_' + dateToday() + '.docx');

  /**
   * Handles file drop event.
   * @param acceptedFiles - The files that were dropped.
   */
  const onDrop = async (acceptedFiles: File[]) => {
    setFiles([...files, ...acceptedFiles]);
  }

  const { getRootProps, getInputProps, isDragActive } = useDropzone({ onDrop });

  /**
   * Handles criteria change event.
   * @param ev - The event object.
   * @param checked - The checked state of the checkbox.
   */
  const handleCriteriaChange = (ev?: FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => {
    const { name } = ev?.target as HTMLInputElement;
    setCriteria(prev => ({ ...prev, [name]: checked }));
  };

  /**
   * Handles output file name change event.
   * @param event - The event object.
   */
  const handleOutputFileNameChange = (event: ChangeEvent<HTMLInputElement>) => {
    setOutputFileName(event.target.value);
  };

  /**
   * Handles save file event.
   */
  const handleSaveFile = async () => {
    const extractedParagraphs: ExtractedParagraph[][] = await Promise.all(files.map(file => extractParagraphs(file, criteria)));
    const fileNames = files.map(file => file.name);
    const sections = buildSections(extractedParagraphs, fileNames);
    const styles = buildStyles();

    const doc = new Document({
      styles,
      sections,
    });

    const buffer = await Packer.toBlob(doc);
    saveAs(buffer, outputFileName);
  };

  /**
   * Moves a file in the list.
   * @param dragIndex - The index of the file being dragged.
   * @param hoverIndex - The index of the file being hovered over.
   */
  const moveFile = (dragIndex: number, hoverIndex: number) => {
    const draggedFile = files[dragIndex];
    const updatedFiles = [...files];
    updatedFiles.splice(dragIndex, 1);
    updatedFiles.splice(hoverIndex, 0, draggedFile);
    setFiles(updatedFiles);
  }

  /**
   * Removes a file from the list.
   * @param index - The index of the file to remove.
   */
  const removeFile = (index: number) => {
    const updatedFiles = files.filter((_, i) => i !== index);
    setFiles(updatedFiles);
  }

  return (
    <DndProvider backend={HTML5Backend}>
      <Stack tokens={{ childrenGap:15, padding: 20}}>
        <div {...getRootProps({ className: `dropzone ${isDragActive ? 'active' : '' }` })}>
          <input {...getInputProps()} />
          <p>Drag 'n' drop some files here, or click to select files</p>
        </div>
        <ul className='file-list'>
          {files.map((file, index) => (
            <FileItem key={file.name} file={file} index={index} moveFile={moveFile} removeFile={removeFile} />
          ))}
        </ul>
        <Stack tokens={{ childrenGap: 10 }}>
          <Checkbox label="Includes redlining" name="redline" checked={criteria.redline} onChange={handleCriteriaChange} />
          <Checkbox label="Includes highlighted text" name="highlight" checked={criteria.highlight} onChange={handleCriteriaChange} />
          <Checkbox label="Includes square brackets" name="squareBrackets" checked={criteria.squareBrackets} onChange={handleCriteriaChange} />
          <Checkbox label="Includes comments" name="comments" checked={criteria.comments} onChange={handleCriteriaChange} />
        </Stack>
        <TextField label="Output File Name" value={outputFileName} onChange={handleOutputFileNameChange} />
        <PrimaryButton text="Save file" onClick={handleSaveFile} />
      </Stack>
    </DndProvider>
  );
};

export default WordHandler;
