import { sp } from "@pnp/sp/presets/all";
import type { FileAttachment } from "../components/approval-form/context/FormContext";

const LIBRARY_NAME = "DocumentosArtes";
const FINAL_APPROVAL_FOLDER = "5.Aprobacion Segunda Etapa";
const FINAL_APPROVAL_FOLDER_PATH =
  `/sites/AprobacionArtes/${LIBRARY_NAME}/${FINAL_APPROVAL_FOLDER}`;

export interface IFinalApprovalFileItem {
  id: number;
  name: string;
  url: string;
  downloadUrl: string;
  modified?: string;
}

interface ILibraryFileListItem {
  Id: number;
  FileLeafRef: string;
  FileRef: string;
  Modified?: string;
}

interface IUploadedLibraryItem {
  Id: number;
  update(data: Record<string, unknown>): Promise<unknown>;
}

export async function uploadFinalApprovalFiles(
  nombreProceso: string,
  procesoId: number,
  files: FileAttachment[]
): Promise<IFinalApprovalFileItem[]> {
  if (files.length === 0) return [];

  const uploadedFiles: IFinalApprovalFileItem[] = [];

  for (const attachment of files) {
    const uploadResult = await sp.web
      .getFolderByServerRelativePath(FINAL_APPROVAL_FOLDER_PATH)
      .files.addUsingPath(attachment.name, attachment.file, { Overwrite: true });

    const uploadedItem = (await sp.web
      .getFileByServerRelativePath(uploadResult.data.ServerRelativeUrl)
      .getItem()) as IUploadedLibraryItem;

    await uploadedItem.update({
      NombreProceso: nombreProceso,
      IDProceso: procesoId,
    });

    const fileUrl = `${window.location.origin}${uploadResult.data.ServerRelativeUrl}`;
    uploadedFiles.push({
      id: uploadedItem.Id,
      name: attachment.name,
      url: fileUrl,
      downloadUrl: `${fileUrl}?download=1`,
    });
  }

  return uploadedFiles;
}

export async function getFinalApprovalFilesByProcess(
  procesoId: number
): Promise<IFinalApprovalFileItem[]> {
  const rows: ILibraryFileListItem[] = await sp.web.lists
    .getByTitle(LIBRARY_NAME)
    .items.select("Id", "FileLeafRef", "FileRef", "Modified")
    .filter(`IDProceso eq ${procesoId}`)
    .orderBy("Modified", false)
    .top(200)
    .get();

  return rows.map((row) => ({
    id: row.Id,
    name: row.FileLeafRef,
    url: `${window.location.origin}${row.FileRef}`,
    downloadUrl: `${window.location.origin}${row.FileRef}?download=1`,
    modified: row.Modified,
  }));
}
