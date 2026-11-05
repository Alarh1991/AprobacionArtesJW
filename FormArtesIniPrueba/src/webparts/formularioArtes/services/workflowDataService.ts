import { sp } from "@pnp/sp/presets/all";
import type { WebPartContext } from "@microsoft/sp-webpart-base";

const LISTA_ETAPAS = "CatalogoEtapas";
const LISTA_APROBADORES = "CatalogoAprobadores";
const LISTA_PRINCIPAL = "1ProcesosAprobacionArtesPrincipal";
const LISTA_PROVEEDORES = "CatalogoProveedores";
const LISTA_MARCAS = "CatalogoMarcas";

interface IEtapaCatalogoItem {
  Id: number;
  CodigoEtapa: string;
  NombreEtapa?: string;
}

interface IAprobadorCatalogoItem {
  Id: number;
  RolAprobador?: string;
  CorreoElectronico?: string;
  EstadoUsuario?: string;
}

interface IProveedorCatalogoItem {
  Id: number;
  Title?: string;
  Proveedor?: string;
  CorreoElectronico?: string;
}

interface ISiteUserItem {
  Id: number;
  Title?: string;
  Email?: string;
}

interface IGraphUserItem {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
}

export interface IProviderOption {
  id: number;
  nombre: string;
  correoElectronico?: string;
}

export interface IUserDirectoryOption {
  id: number | string;
  nombre: string;
  email: string;
}

interface IMarcaCatalogoItem {
  Id: number;
  Title?: string;
  NombreMarca?: string;
}

export interface IBrandOption {
  id: number;
  nombre: string;
}

export function normalizeEtapaCode(value: string): string {
  return value.trim().toUpperCase();
}

/* =========================================================
   CatalogoEtapas: CodigoEtapa -> Id
========================================================= */
export async function getEtapaIdByCodeMap(): Promise<Record<string, number>> {
  const etapas: IEtapaCatalogoItem[] = await sp.web.lists
    .getByTitle(LISTA_ETAPAS)
    .items.select("Id", "CodigoEtapa", "NombreEtapa")
    .get();

  const map: Record<string, number> = {};

  for (const etapa of etapas) {
    if (etapa.CodigoEtapa) {
      map[normalizeEtapaCode(etapa.CodigoEtapa)] = etapa.Id;
    }

    // Fallback por si en algún ambiente el registro vino con código mal cargado.
    if (
      etapa.NombreEtapa?.toLowerCase().includes("rechazado inicial regulatorio")
    ) {
      map["R2.3"] = etapa.Id;
    }
  }

  return map;
}

export async function resolveEtapaIdOrThrow(codigoEtapa: string): Promise<number> {
  const map = await getEtapaIdByCodeMap();
  const etapaId = map[normalizeEtapaCode(codigoEtapa)];

  if (!etapaId) {
    throw new Error(
      `No se encontró EtapaId para el código '${codigoEtapa}' en '${LISTA_ETAPAS}'.`
    );
  }

  return etapaId;
}

/* =========================================================
   CatalogoAprobadores: rol del usuario actual
========================================================= */
export async function getCurrentUserRole(): Promise<string | undefined> {
  const roles = await getCurrentUserRoles();
  return roles[0];
}

export async function getCurrentUserRoles(): Promise<string[]> {
  const user = await sp.web.currentUser();

  const rows: IAprobadorCatalogoItem[] = await sp.web.lists
    .getByTitle(LISTA_APROBADORES)
    .items.select("Id", "RolAprobador", "CorreoElectronico", "EstadoUsuario")
    .filter(`CorreoElectronico eq '${user.Email}' and EstadoUsuario eq 'Activo'`)
    .top(5000)
    .get();

  const uniqueRoles = Array.from(
    new Set(
      rows
        .map((row) => row.RolAprobador?.trim())
        .filter((rol): rol is string => Boolean(rol))
    )
  );

  return uniqueRoles;
}

/* =========================================================
   Lista principal: utilidades de lectura
========================================================= */
export async function getProcessDocumentLink(procesoId: number): Promise<string> {
  const principal = await sp.web.lists
    .getByTitle(LISTA_PRINCIPAL)
    .items.getById(procesoId)
    .select("URLDocumentos")
    .get();

  return principal?.URLDocumentos?.Url ?? "";
}

/* =========================================================
   CatalogoProveedores: opciones para Step3
========================================================= */
export async function getProviderOptions(): Promise<IProviderOption[]> {
  const providers: IProveedorCatalogoItem[] = await sp.web.lists
    .getByTitle(LISTA_PROVEEDORES)
    .items.select("Id", "Title", "Proveedor", "CorreoElectronico")
    .top(5000)
    .get();

  return providers
    .map((row) => ({
      id: row.Id,
      nombre: (row.Proveedor ?? row.Title ?? "").trim(),
      correoElectronico: row.CorreoElectronico?.trim(),
    }))
    .filter((row) => row.nombre.length > 0)
    .sort((a, b) => a.nombre.localeCompare(b.nombre, "es", { sensitivity: "base" }));
}

/* =========================================================
   CatalogoMarcas: opciones para Step1
========================================================= */
export async function getBrandOptions(): Promise<IBrandOption[]> {
  const marcas: IMarcaCatalogoItem[] = await sp.web.lists
    .getByTitle(LISTA_MARCAS)
    .items.select("Id", "Title", "NombreMarca")
    .top(5000)
    .get();

  return marcas
    .map((row) => ({
      id: row.Id,
      nombre: (row.NombreMarca ?? row.Title ?? "").trim(),
    }))
    .filter((row) => row.nombre.length > 0)
    .sort((a, b) => a.nombre.localeCompare(b.nombre, "es", { sensitivity: "base" }));
}

/* =========================================================
   Usuarios del sitio: busqueda por nombre o correo para copias
========================================================= */
export async function searchSiteUserOptions(
  searchText: string
): Promise<IUserDirectoryOption[]> {
  const normalizedText = searchText.trim();
  if (normalizedText.length < 2) return [];

  const escapedText = normalizedText.replace(/'/g, "''");
  const users: ISiteUserItem[] = await sp.web.siteUsers
    .select("Id", "Title", "Email")
    .filter(
      `substringof('${escapedText}', Title) or substringof('${escapedText}', Email)`
    )
    .top(10)
    .get();

  return users
    .map((user) => ({
      id: user.Id,
      nombre: (user.Title ?? "").trim(),
      email: (user.Email ?? "").trim(),
    }))
    .filter((user) => user.email.length > 0)
    .sort((a, b) => a.nombre.localeCompare(b.nombre, "es", { sensitivity: "base" }));
}

export async function searchDirectoryUserOptions(
  spfxContext: WebPartContext,
  searchText: string
): Promise<IUserDirectoryOption[]> {
  const normalizedText = searchText.trim();
  if (normalizedText.length < 2) return [];

  const escapedText = normalizedText.replace(/'/g, "''");
  const client = await spfxContext.msGraphClientFactory.getClient("3");
  const response = await client
    .api("/users")
    .version("v1.0")
    .header("ConsistencyLevel", "eventual")
    .filter(
      `startsWith(displayName,'${escapedText}') or startsWith(mail,'${escapedText}') or startsWith(userPrincipalName,'${escapedText}')`
    )
    .select("id,displayName,mail,userPrincipalName")
    .top(12)
    .get();

  const users: IGraphUserItem[] = Array.isArray(response?.value) ? response.value : [];

  return users
    .map((user) => ({
      id: user.id,
      nombre: (user.displayName ?? user.mail ?? user.userPrincipalName ?? "").trim(),
      email: (user.mail ?? user.userPrincipalName ?? "").trim(),
    }))
    .filter((user) => user.nombre.length > 0 && user.email.length > 0)
    .sort((a, b) => a.nombre.localeCompare(b.nombre, "es", { sensitivity: "base" }));
}
