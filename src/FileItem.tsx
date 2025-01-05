// @ts-ignore
import React from 'react';
import { useDrag, useDrop, DropTargetMonitor } from 'react-dnd';
import { IconButton } from '@fluentui/react';

/**
 * Props for the FileItem component.
 */
interface FileItemProps {
  file: File;
  index: number;
  moveFile: (dragIndex: number, hoverIndex: number) => void;
  removeFile: (index: number) => void;
}

/**
 * Component representing a file item in the list.
 * @param file - The file object.
 * @param index - The index of the file in the list.
 * @param moveFile - Function to move the file in the list.
 * @param removeFile - Function to remove the file from the list.
 * @returns A React component.
 */
const FileItem: React.FC<FileItemProps> = ({ file, index, moveFile, removeFile }) => {
  const ref = React.useRef<HTMLLIElement>(null);

  const [, drop] = useDrop({
    accept: 'file',
    hover(item: {index: number }, monitor: DropTargetMonitor) {
      if (!ref.current) {
        return;
      }
      const dragIndex = item.index;
      const hoverIndex = index;

      if (dragIndex === hoverIndex) {
        return;
      }

      const hoverBoundingRect = ref.current.getBoundingClientRect();
      const hoverMiddleY = (hoverBoundingRect.bottom - hoverBoundingRect.top) / 2;
      const clientOffset = monitor.getClientOffset();
      const hoverClientY = clientOffset.y - hoverBoundingRect.top;

      if (dragIndex < hoverIndex && hoverClientY < hoverMiddleY) {
        return;
      }

      if (dragIndex > hoverIndex && hoverClientY > hoverMiddleY) {
        return;
      }

      moveFile(dragIndex, hoverIndex);
      item.index = hoverIndex;
    },
  });

  const [, drag] = useDrag({
    type: 'file',
    item: { type: file, index },
  });

  drag(drop(ref));

  return (
    <li ref={ref} className="filename">
      {file.name}
      <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => removeFile(index)} />
    </li>
  );
};

export default FileItem;