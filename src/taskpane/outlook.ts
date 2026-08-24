/* global document, Office */

import {
  createNestablePublicClientApplication,
  InteractionRequiredAuthError,
  type AccountInfo,
  type IPublicClientApplication,
} from "@azure/msal-browser";

const clientId = "5765f492-946f-4ef2-a8d6-416514b589a6";
const tenantId = "c67c1288-4131-4879-8c88-ae0bc631308c";
const scopes = ["User.Read", "Sites.ReadWrite.All"];

let authClient: IPublicClientApplication;

interface GraphCollection<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

interface SiteResult {
  id: string;
}

interface ListResult {
  id: string;
  displayName: string;
}

interface ColumnResult {
  name: string;
  displayName: string;
}

interface ListItemResult {
  id: string;
  fields: Record<string, unknown>;
}

interface DriveResult {
  id: string;
  name: string;
}

interface DriveItemResult {
  id: string;
  name: string;
  webUrl?: string;
  folder?: { childCount: number };
}

let currentAccessToken = "";
let documentsDriveId = "";
let selectedEmailAttachmentCount = 0;

async function getAuthClient(): Promise<IPublicClientApplication> {
  authClient ??= await createNestablePublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      redirectUri: "brk-multihub://localhost:3000",
    },
    cache: { cacheLocation: "localStorage" },
  });
  return authClient;
}

async function graphGet<T>(url: string, accessToken: string): Promise<T> {
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!response.ok) {
    const details = await response.text();
    const error = new Error(`Microsoft Graph returned ${response.status}: ${details}`) as Error & {
      status?: number;
    };
    error.status = response.status;
    throw error;
  }

  return response.json() as Promise<T>;
}

async function graphJson<T>(url: string, method: string, body: unknown, accessToken: string): Promise<T> {
  const response = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const details = await response.text();
    const error = new Error(`Microsoft Graph returned ${response.status}: ${details}`) as Error & {
      status?: number;
    };
    error.status = response.status;
    throw error;
  }

  return response.json() as Promise<T>;
}

async function freshAccessToken(): Promise<string> {
  authClient = await getAuthClient();
  const account = authClient.getAllAccounts()[0];
  if (!account) throw new Error("Microsoft 365 is not connected.");
  return (await authClient.acquireTokenSilent({ account, scopes })).accessToken;
}

async function loadActiveProjects(accessToken: string): Promise<void> {
  const select = document.getElementById("project-select") as HTMLSelectElement;
  const status = document.getElementById("status")!;
  status.textContent = "Loading active projects…";

  const site = await graphGet<SiteResult>(
    "https://graph.microsoft.com/v1.0/sites/netorg14705176.sharepoint.com:/sites/LESARCHITECTS-PROJECTS",
    accessToken
  );

  const lists = await graphGet<GraphCollection<ListResult>>(
    `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id)}/lists?$select=id,displayName`,
    accessToken
  );
  const projectList = lists.value.find((list) => list.displayName === "Project Directory");
  if (!projectList) throw new Error("The SharePoint Project Directory list was not found.");

  const drives = await graphGet<GraphCollection<DriveResult>>(
    `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id)}/drives?$select=id,name`,
    accessToken
  );
  const documentsDrive = drives.value.find((drive) =>
    ["documents", "shared documents"].includes(drive.name.toLowerCase())
  );
  if (!documentsDrive) throw new Error("The SharePoint Documents library was not found.");
  documentsDriveId = documentsDrive.id;

  const columns = await graphGet<GraphCollection<ColumnResult>>(
    `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id)}/lists/${projectList.id}/columns?$select=name,displayName`,
    accessToken
  );
  const columnName = (displayName: string) => {
    const column = columns.value.find((candidate) => candidate.displayName === displayName);
    if (!column) throw new Error(`The SharePoint column '${displayName}' was not found.`);
    return column.name;
  };

  const projectNoField = columnName("Project No");
  const titleField = columnName("Title");
  const statusField = columnName("Status");
  let itemsUrl = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id)}/lists/${projectList.id}/items?$expand=fields`;
  const items: ListItemResult[] = [];

  while (itemsUrl) {
    const page = await graphGet<GraphCollection<ListItemResult>>(itemsUrl, accessToken);
    items.push(...page.value);
    itemsUrl = page["@odata.nextLink"] ?? "";
  }

  const projects = items
    .filter((item) => String(item.fields[statusField] ?? "").toLowerCase() !== "completed")
    .map((item) => ({
      id: item.id,
      projectNo: String(item.fields[projectNoField] ?? "").trim(),
      name: String(item.fields[titleField] ?? "Untitled project").trim(),
    }))
    .sort((a, b) => a.projectNo.localeCompare(b.projectNo, undefined, { numeric: true }));

  select.innerHTML = '<option value="">Select a project…</option>';
  for (const project of projects) {
    const option = document.createElement("option");
    option.value = project.id;
    option.dataset.projectNo = project.projectNo;
    option.textContent = `${project.projectNo} — ${project.name}`;
    select.appendChild(option);
  }
  select.disabled = false;
  status.textContent = `${projects.length} active projects loaded.`;
}

async function getChildFolders(parentItemId?: string): Promise<DriveItemResult[]> {
  const endpoint = parentItemId
    ? `https://graph.microsoft.com/v1.0/drives/${documentsDriveId}/items/${parentItemId}/children?$select=id,name,folder`
    : `https://graph.microsoft.com/v1.0/drives/${documentsDriveId}/root/children?$select=id,name,folder`;
  const result = await graphGet<GraphCollection<DriveItemResult>>(endpoint, currentAccessToken);
  return result.value.filter((item) => item.folder);
}

async function findChildFolder(parentItemId: string, folderName: string): Promise<DriveItemResult> {
  const folders = await getChildFolders(parentItemId);
  const match = folders.find((folder) => folder.name.toLowerCase() === folderName.toLowerCase());
  if (!match) throw new Error(`The folder '${folderName}' was not found.`);
  return match;
}

async function loadDestinationFolders(): Promise<void> {
  const projectSelect = document.getElementById("project-select") as HTMLSelectElement;
  const folderField = document.getElementById("folder-field")!;
  const folderSelect = document.getElementById("folder-select") as HTMLSelectElement;
  const saveButton = document.getElementById("save-button") as HTMLButtonElement;
  const status = document.getElementById("status")!;
  const selectedOption = projectSelect.selectedOptions[0];

  folderField.hidden = !projectSelect.value;
  folderSelect.disabled = true;
  saveButton.disabled = true;
  if (!projectSelect.value) return;

  folderSelect.innerHTML = "<option>Loading folders…</option>";
  status.textContent = "Loading destination folders…";

  try {
    const projectNo = selectedOption.dataset.projectNo!;
    const rootFolders = await getChildFolders();
    const prefix = `${projectNo}_`.toLowerCase();
    const projectFolder = rootFolders.find((folder) => folder.name.toLowerCase().startsWith(prefix));
    if (!projectFolder) throw new Error(`No Documents folder starts with '${projectNo}_'.`);

    const drawingsFolder = await findChildFolder(projectFolder.id, "02_DRAWINGS");
    const filesReceivedFolder = await findChildFolder(drawingsFolder.id, "02_FILES RECEIVED");
    const destinationFolders = (await getChildFolders(filesReceivedFolder.id)).sort((a, b) =>
      a.name.localeCompare(b.name, undefined, { numeric: true })
    );

    folderSelect.innerHTML = '<option value="">Select a folder…</option>';
    for (const folder of destinationFolders) {
      const option = document.createElement("option");
      option.value = folder.id;
      option.textContent = folder.name;
      folderSelect.appendChild(option);
    }
    folderSelect.disabled = false;
    status.textContent = `${destinationFolders.length} destination folders loaded.`;
  } catch (error) {
    folderSelect.innerHTML = "<option>Folder structure unavailable</option>";
    status.textContent = error instanceof Error ? error.message : "Destination folders could not be loaded.";
  }
}

function buildFolderName(): string {
  const descriptionInput = document.getElementById("folder-description") as HTMLInputElement;
  const now = new Date();
  const datePrefix = [
    String(now.getFullYear()).slice(-2),
    String(now.getMonth() + 1).padStart(2, "0"),
    String(now.getDate()).padStart(2, "0"),
  ].join("");
  const safeDescription = descriptionInput.value
    .trim()
    .replace(/["*:<>?/\\|#%]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/[. ]+$/g, "");

  return safeDescription ? `${datePrefix}_${safeDescription}` : "";
}

function updateSaveButton(): void {
  const folderSelect = document.getElementById("folder-select") as HTMLSelectElement;
  const descriptionField = document.getElementById("description-field")!;
  const descriptionInput = document.getElementById("folder-description") as HTMLInputElement;
  const preview = document.getElementById("folder-name-preview")!;
  const saveButton = document.getElementById("save-button") as HTMLButtonElement;
  descriptionField.hidden = !folderSelect.value;

  const folderName = buildFolderName();
  const datePrefix = [String(new Date().getFullYear()).slice(-2), String(new Date().getMonth() + 1).padStart(2, "0"), String(new Date().getDate()).padStart(2, "0")].join("");
  preview.textContent = folderName || `${datePrefix}_…`;
  saveButton.disabled = !folderSelect.value || !folderName || selectedEmailAttachmentCount === 0;
}

function getAttachmentContent(attachmentId: string): Promise<Office.AttachmentContent> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error.message));
    });
  });
}

function base64Blob(content: string, contentType: string): Blob {
  const binary = atob(content);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return new Blob([bytes], { type: contentType || "application/octet-stream" });
}

function safeFileName(name: string): string {
  return name.replace(/["*:<>?/\\|#%]/g, "-").replace(/[. ]+$/g, "") || "attachment";
}

async function saveAttachments(): Promise<void> {
  const saveButton = document.getElementById("save-button") as HTMLButtonElement;
  const folderSelect = document.getElementById("folder-select") as HTMLSelectElement;
  const status = document.getElementById("status")!;
  const folderName = buildFolderName();
  const attachments = Office.context.mailbox.item.attachments.filter((attachment) => !attachment.isInline);

  if (!folderSelect.value || !folderName || attachments.length === 0) return;

  saveButton.disabled = true;
  status.textContent = `Creating ${folderName}…`;

  try {
    currentAccessToken = await freshAccessToken();
    const createdFolder = await graphJson<DriveItemResult>(
      `https://graph.microsoft.com/v1.0/drives/${documentsDriveId}/items/${folderSelect.value}/children`,
      "POST",
      { name: folderName, folder: {}, "@microsoft.graph.conflictBehavior": "fail" },
      currentAccessToken
    );

    const usedNames = new Map<string, number>();
    for (let index = 0; index < attachments.length; index += 1) {
      const attachment = attachments[index];
      const originalName = safeFileName(attachment.name);
      const key = originalName.toLowerCase();
      const occurrence = (usedNames.get(key) ?? 0) + 1;
      usedNames.set(key, occurrence);
      const dot = originalName.lastIndexOf(".");
      const fileName = occurrence === 1
        ? originalName
        : dot > 0
          ? `${originalName.slice(0, dot)} (${occurrence})${originalName.slice(dot)}`
          : `${originalName} (${occurrence})`;

      status.textContent = `Uploading ${index + 1} of ${attachments.length}: ${fileName}`;
      const content = await getAttachmentContent(attachment.id);
      if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
        throw new Error(`'${attachment.name}' isn't a downloadable file attachment.`);
      }

      const upload = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${documentsDriveId}/items/${createdFolder.id}:/${encodeURIComponent(fileName)}:/content`,
        {
          method: "PUT",
          headers: { Authorization: `Bearer ${currentAccessToken}` },
          body: base64Blob(content.content, attachment.contentType),
        }
      );
      if (!upload.ok) throw new Error(`Uploading '${fileName}' failed with status ${upload.status}.`);
    }

    status.textContent = `Saved ${attachments.length} attachment${attachments.length === 1 ? "" : "s"}. `;
    if (createdFolder.webUrl) {
      const link = document.createElement("a");
      link.href = createdFolder.webUrl;
      link.target = "_blank";
      link.rel = "noopener";
      link.textContent = "Open folder";
      status.appendChild(link);
    }
    saveButton.textContent = "Attachments saved";
  } catch (error) {
    if (error instanceof Error && (error as Error & { status?: number }).status === 409) {
      status.textContent = `A folder named '${folderName}' already exists. Enter a different folder description.`;
    } else {
      status.textContent = error instanceof Error ? error.message : "The attachments could not be saved.";
    }
    saveButton.disabled = false;
  }
}

async function connectMicrosoft365(): Promise<void> {
  const button = document.getElementById("connect-button") as HTMLButtonElement;
  const status = document.getElementById("status")!;

  button.disabled = true;
  status.textContent = "Connecting…";

  try {
    authClient = await getAuthClient();

    let account: AccountInfo | undefined = authClient.getAllAccounts()[0];

    if (!account) {
      const login = await authClient.loginPopup({ scopes });
      account = login.account ?? undefined;
    }

    if (!account) throw new Error("Microsoft 365 did not return a signed-in account.");

    let accessToken: string;
    try {
      accessToken = (await authClient.acquireTokenSilent({ account, scopes })).accessToken;
    } catch (error) {
      if (!(error instanceof InteractionRequiredAuthError)) throw error;
      accessToken = (await authClient.acquireTokenPopup({ account, scopes })).accessToken;
    }

    button.hidden = true;
    currentAccessToken = accessToken;
    await loadActiveProjects(accessToken);
  } catch (error) {
    button.disabled = false;
    status.textContent = error instanceof Error ? error.message : "Microsoft 365 connection failed.";
  }
}

async function connectSilently(): Promise<void> {
  const button = document.getElementById("connect-button") as HTMLButtonElement;
  const status = document.getElementById("status")!;

  try {
    authClient = await getAuthClient();
    const account = authClient.getAllAccounts()[0];
    if (!account) return;

    status.textContent = "Connecting to Microsoft 365…";
    const token = await authClient.acquireTokenSilent({ account, scopes });
    currentAccessToken = token.accessToken;
    button.hidden = true;
    await loadActiveProjects(token.accessToken);
  } catch {
    button.hidden = false;
    status.textContent = "Select Connect Microsoft 365 to continue.";
  }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  const appBody = document.getElementById("app-body");
  const subject = document.getElementById("email-subject");
  const attachmentCount = document.getElementById("attachment-count");
  const connectButton = document.getElementById("connect-button");
  const projectSelect = document.getElementById("project-select");
  const folderSelect = document.getElementById("folder-select");
  const descriptionInput = document.getElementById("folder-description");
  const saveButton = document.getElementById("save-button");
  const item = Office.context.mailbox.item;

  if (!appBody || !subject || !attachmentCount || !connectButton || !projectSelect || !folderSelect || !descriptionInput || !saveButton || !item) return;

  subject.textContent = item.subject || "(No subject)";
  selectedEmailAttachmentCount = item.attachments?.filter((attachment) => !attachment.isInline).length ?? 0;
  attachmentCount.textContent = `${selectedEmailAttachmentCount} file attachment${selectedEmailAttachmentCount === 1 ? "" : "s"}`;
  connectButton.addEventListener("click", connectMicrosoft365);
  projectSelect.addEventListener("change", loadDestinationFolders);
  folderSelect.addEventListener("change", updateSaveButton);
  descriptionInput.addEventListener("input", updateSaveButton);
  saveButton.addEventListener("click", saveAttachments);
  appBody.style.display = "block";
  void connectSilently();
});
