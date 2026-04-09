import React from 'react';
import { DndProvider } from 'react-dnd';
import { HTML5Backend } from 'react-dnd-html5-backend';
import { initializeIcons } from '@fluentui/font-icons-mdl2';
import { Checkbox, Stack, Label, PrimaryButton, TextField, } from '@fluentui/react';
import { useDropzone } from 'react-dropzone';

import FileItem from './FileItem';
import { criteriaOptions } from './wordHandlerConfig';
import { useWordHandler } from './useWordHandler';

initializeIcons(); // For checkmark icon

/**
 * Main component for handling Word document processing.
 * @returns A React component.
 */
const WordHandler: React.FC = () => {
  const {
    files,
    criteria,
    outputFileName,
    addFiles,
    handleCriteriaChange,
    handleOutputFileNameChange,
    handleSaveFile,
    moveFile,
    removeFile,
  } = useWordHandler();

  const { getRootProps, getInputProps, isDragActive } = useDropzone({ onDrop: addFiles });

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
          {criteriaOptions.map(({ key, label }) => (
            <Checkbox
              key={key}
              label={label}
              name={key}
              checked={criteria[key]}
              onChange={handleCriteriaChange}
            />
          ))}
        </Stack>
        <Label>Output File Name</Label>
        <TextField value={outputFileName} onChange={handleOutputFileNameChange} />
        <PrimaryButton text="Save file" onClick={handleSaveFile} />
      </Stack>
    </DndProvider>
  );
};

export default WordHandler;
