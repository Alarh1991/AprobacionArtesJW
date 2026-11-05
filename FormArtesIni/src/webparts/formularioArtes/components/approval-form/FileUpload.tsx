import * as React from "react";
import { useRef, useEffect, useState } from "react";

import {
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
} from "@fluentui/react";

import styles from "./styles/form.module.scss";

/* ================================
   MODELO DE ADJUNTO
   ================================ */
export interface FileAttachment {
  id: string;
  name: string;
  size: number;
  type: string;
  file: File;
  relativePath?: string; // SOLO para carpetas
}

interface IFileWithRelativePath extends File {
  readonly webkitRelativePath: string;
}

/* ================================
   PROPS
   ================================ */
interface FileUploadProps {
  files: FileAttachment[];
  onFilesChange: (files: FileAttachment[]) => void;
  disabled?: boolean;
  allowFolders?: boolean;
  openFileDialogOnClick?: boolean;
  uploadText?: string;
  helperText?: string;
}

/* ================================
   COMPONENTE
   ================================ */
export const FileUpload: React.FC<FileUploadProps> = ({
  files,
  onFilesChange,
  disabled = false,
  allowFolders = true,
  openFileDialogOnClick = false,
  uploadText = "Haz clic para agregar archivos o carpetas",
  helperText = "Puedes subir archivos individuales o carpetas completas",
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);

  const [showDialog, setShowDialog] = useState(false);

  /* ================================
     ACTIVAR MODO CARPETA
     ================================ */
  useEffect(() => {
    if (folderInputRef.current) {
      folderInputRef.current.setAttribute("webkitdirectory", "");
      folderInputRef.current.setAttribute("directory", "");
    }
  }, []);

  /* ================================
     MANEJO DE ARCHIVOS
     ================================ */
  const processFiles = (fileList: FileList | null): void => {
    if (!fileList) return;

    const newFiles: FileAttachment[] = Array.from(fileList).map((file) => ({
      id: `${Date.now()}-${Math.random().toString(36).substring(2, 9)}`,
      name: file.name,
      size: file.size,
      type: file.type,
      file,
      relativePath: (file as IFileWithRelativePath).webkitRelativePath || undefined,
    }));

    onFilesChange([...files, ...newFiles]);
  };

  /* ================================
     EVENTOS
     ================================ */
  const handleFilesSelected = (e: React.ChangeEvent<HTMLInputElement>): void => {
    processFiles(e.target.files);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const handleFolderSelected = (e: React.ChangeEvent<HTMLInputElement>): void => {
    processFiles(e.target.files);
    if (folderInputRef.current) folderInputRef.current.value = "";
  };

  const handleRemove = (id: string): void => {
    onFilesChange(files.filter((f) => f.id !== id));
  };

  /* ================================
     UTILIDADES
     ================================ */
  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return "0 Bytes";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return `${(bytes / Math.pow(k, i)).toFixed(2)} ${sizes[i]}`;
  };

  const getIcon = (file: FileAttachment): string => {
    if (file.relativePath) return "📁";
    if (file.type.startsWith("image/")) return "🖼️";
    if (file.type.includes("pdf")) return "📄";
    if (file.type.includes("word")) return "📝";
    if (file.type.includes("excel")) return "📊";
    return "📎";
  };

  /* ================================
     RENDER
     ================================ */
  return (
    <>
      {/* ================= MODAL ================= */}
      <Dialog
        hidden={!showDialog || !allowFolders}
        onDismiss={() => setShowDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "¿Qué deseas subir?",
          subText: "Selecciona el tipo de elemento que deseas adjuntar",
        }}
      >
        <DialogFooter>
          <PrimaryButton
            text="📄 Archivos sueltos"
            onClick={() => {
              setShowDialog(false);
              fileInputRef.current?.click();
            }}
          />
          {allowFolders && (
            <DefaultButton
              text="📁 Carpeta completa"
              onClick={() => {
                setShowDialog(false);
                folderInputRef.current?.click();
              }}
            />
          )}
        </DialogFooter>
      </Dialog>

      {/* ================= ZONA DE CARGA ================= */}
      <div
        className={styles.fileUpload}
        onClick={() => {
          if (disabled) return;

          if (openFileDialogOnClick || !allowFolders) {
            fileInputRef.current?.click();
            return;
          }

          setShowDialog(true);
        }}
        style={{
          opacity: disabled ? 0.6 : 1,
          cursor: disabled ? "not-allowed" : "pointer",
        }}
      >
        <div className={styles.fileUploadIcon}>📁</div>
        <p className={styles.fileUploadText}>{uploadText}</p>
        <p className={styles.helperText}>{helperText}</p>
      </div>

      {/* ================= INPUTS OCULTOS ================= */}
      <input
        ref={fileInputRef}
        type="file"
        multiple
        style={{ display: "none" }}
        onChange={handleFilesSelected}
        disabled={disabled}
      />

      {allowFolders && (
        <input
          ref={folderInputRef}
          type="file"
          multiple
          style={{ display: "none" }}
          onChange={handleFolderSelected}
          disabled={disabled}
        />
      )}

      {/* ================= LISTADO ================= */}
      {files.length > 0 && (
        <ul className={styles.fileList}>
          {files.map((file) => (
            <li key={file.id} className={styles.fileItem}>
              <span className={styles.fileName}>
                {getIcon(file)}{" "}
                {file.relativePath
                  ? file.relativePath
                  : file.name}
              </span>

              <span className={styles.fileSize}>
                {formatFileSize(file.size)}
                {!disabled && (
                  <button
                    type="button"
                    className={styles.fileRemove}
                    onClick={() => handleRemove(file.id)}
                  >
                    ×
                  </button>
                )}
              </span>
            </li>
          ))}
        </ul>
      )}
    </>
  );
};
