import { FormEvent, useState } from 'react';
import { saveAs } from 'file-saver';
import { Document, Packer } from 'docx';
import { buildSections, buildStyles, extractParagraphs } from './wordUtils';
import { Criteria } from './types';
import { defaultCriteria, getDefaultOutputFileName } from './wordHandlerConfig';

const moveItem = <T,>(items: T[], fromIndex: number, toIndex: number): T[] => {
  const nextItems = [...items];
  const [movedItem] = nextItems.splice(fromIndex, 1);
  nextItems.splice(toIndex, 0, movedItem);
  return nextItems;
};

export const useWordHandler = () => {
  const [files, setFiles] = useState<File[]>([]);
  const [criteria, setCriteria] = useState<Criteria>(defaultCriteria);
  const [outputFileName, setOutputFileName] = useState<string>(getDefaultOutputFileName);

  const addFiles = (acceptedFiles: File[]) => {
    setFiles(currentFiles => [...currentFiles, ...acceptedFiles]);
  };

  const handleCriteriaChange = (ev?: FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => {
    const { name } = ev?.target as HTMLInputElement;
    setCriteria(currentCriteria => ({ ...currentCriteria, [name]: checked ?? false }));
  };

  const handleOutputFileNameChange = (_ev?: FormEvent<HTMLElement | HTMLInputElement>, newValue?: string) => {
    setOutputFileName(newValue || '');
  };

  const moveFile = (dragIndex: number, hoverIndex: number) => {
    setFiles(currentFiles => moveItem(currentFiles, dragIndex, hoverIndex));
  };

  const removeFile = (index: number) => {
    setFiles(currentFiles => currentFiles.filter((_, currentIndex) => currentIndex !== index));
  };

  const handleSaveFile = async () => {
    const extractedParagraphs = await Promise.all(files.map(file => extractParagraphs(file, criteria)));
    const fileNames = files.map(file => file.name);
    const sections = buildSections(extractedParagraphs, fileNames);
    const styles = buildStyles();
    const doc = new Document({ styles, sections });
    const buffer = await Packer.toBlob(doc);

    saveAs(buffer, outputFileName);
  };

  return {
    files,
    criteria,
    outputFileName,
    addFiles,
    handleCriteriaChange,
    handleOutputFileNameChange,
    handleSaveFile,
    moveFile,
    removeFile,
  };
};
