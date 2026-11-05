import { sp } from "@pnp/sp";
import type { FileAttachment } from
  "../components/approval-form/context/FormContext";

const BASE_LIBRARY_PATH =
  "/sites/AprobacionArtes/DocumentosArtes/1.TARJETAS IMPRESIÓN INTERNAS";

export async function uploadAttachments(
  marca: string,
  requestFolder: string,
  files: FileAttachment[]
): Promise<void> {

  const marcaPath = `${BASE_LIBRARY_PATH}/${marca}`;
  const procesoPath = `${marcaPath}/${requestFolder}`;

  console.log("📂 Ruta proceso:", procesoPath);

  // 1️⃣ Validar / crear carpeta marca
  try {
    await sp.web.getFolderByServerRelativePath(marcaPath).get();
  } catch {
    await sp.web.folders.addUsingPath(marcaPath);
  }

  // 2️⃣ Validar / crear carpeta proceso
  try {
    await sp.web.getFolderByServerRelativePath(procesoPath).get();
  } catch {
    await sp.web.folders.addUsingPath(procesoPath);
  }

  // 3️⃣ Subir archivos
  for (const item of files) {

    let targetFolder = procesoPath;

    if (item.relativePath) {
      const parts = item.relativePath.split("/");
      const folders = parts.slice(0, -1);

      for (const folder of folders) {
        targetFolder += `/${folder}`;
        try {
          await sp.web.getFolderByServerRelativePath(targetFolder).get();
        } catch {
          await sp.web.folders.addUsingPath(targetFolder);
        }
      }
    }

    await sp.web
      .getFolderByServerRelativePath(targetFolder)
      .files.addUsingPath(item.name, item.file, { Overwrite: true });

    console.log("✅ Subido:", `${targetFolder}/${item.name}`);
  }
}