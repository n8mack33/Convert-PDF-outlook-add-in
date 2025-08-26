// import * as React from 'react';
// import {
//   Stack,
//   Text,
//   PrimaryButton,
//   Checkbox,
//   TextField,
//   Spinner,
//   SpinnerSize,
//   MessageBar,
//   MessageBarType,
//   initializeIcons
// } from '@fluentui/react';
// import { useBoolean } from '@fluentui/react-hooks';
// import { Get_Token_SSO } from './SSO_For_Graph'; // Yeh file aek valid Graph token faraham karti hai

// // Fluent UI icons ko initialize karein
// initializeIcons();

// // --- CONSTANTS ---
// const MAX_FILENAME_LENGTH = 250; // Attachment names ke liye aek mehfooz length
// const MAX_FILE_SIZE_MB = 10;
// const GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0";
// const TEMP_ONEDRIVE_FOLDER = "TempAddinUploads"; // Temporary files ke liye OneDrive folder

// // --- TYPES & INTERFACES ---
// type AttachmentStatus = 'pending' | 'processing' | 'success' | 'error';
// type GlobalStatusType = 'info' | 'success' | 'error' | 'warning';

// interface IAttachment {
//   id: string;
//   name: string;
//   newName: string;
//   isSelected: boolean;
//   contentType: string;
//   status: AttachmentStatus;
//   statusMessage: string;
// }

// interface IGlobalStatus {
//   message: string;
//   type: GlobalStatusType;
// }

// // --- LOGGING UTILITY ---
// const log = (message: string, data?: any) => {
//   console.log(`[PDF Add-in] ${message}`, data !== undefined ? data : '');
// };

// // --- OFFICE.JS ASYNC WRAPPERS (Modern async/await pattern ke liye) ---
// const getAttachmentsAsync = (): Promise<Office.AttachmentDetails[]> => {
//   return new Promise((resolve, reject) => {
//     Office.context.mailbox.item.getAttachmentsAsync((result:any) => {
//       if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
//       else reject(new Error(result.error.message));
//     });
//   });
// };

// const getAttachmentContentAsync = (id: string): Promise<ArrayBuffer> => {
//   return new Promise((resolve, reject) => {
//     Office.context.mailbox.item.getAttachmentContentAsync(id, (result) => {
//       if (result.status === Office.AsyncResultStatus.Succeeded) {
//         try {
//           const binaryString = window.atob(result.value.content);
//           const bytes = new Uint8Array(binaryString.length);
//           for (let i = 0; i < binaryString.length; i++) {
//             bytes[i] = binaryString.charCodeAt(i);
//           }
//           resolve(bytes.buffer);
//         } catch (error) {
//           reject(new Error(`Base64 content decode karne mein naakam: ${error.message}`));
//         }
//       } else {
//         reject(new Error(result.error.message));
//       }
//     });
//   });
// };

// const addFileAttachmentFromBase64Async = (
//   base64: string,
//   name: string
// ): Promise<string> => {
//   return new Promise((resolve, reject) => {
//     Office.context.mailbox.item.addFileAttachmentFromBase64Async(
//       base64,
//       name,
//       { isInline: false },
//       (asyncResult) => {
//         if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
//           resolve(asyncResult.value); // this is the attachmentId
//         } else {
//           reject(
//             new Error(
//               `Attachment failed. Code: ${asyncResult.error.code}, Message: ${asyncResult.error.message}`
//             )
//           );
//         }
//       }
//     );
//   });
// };


// const removeAttachmentAsync = (id: string): Promise<void> => {
//   return new Promise((resolve) => {
//     Office.context.mailbox.item.removeAttachmentAsync(id, (result) => {
//       if (result.status !== Office.AsyncResultStatus.Succeeded) {
//         console.warn(`Asal attachment nahi hata, lekin process jaari hai. Error: ${result.error.message}`);
//       }
//       resolve(); // Hamesha resolve karein taake process chalta rahe.
//     });
//   });
// };

// const saveItemAsync = (): Promise<void> => {
//   return new Promise((resolve, reject) => {
//     Office.context.mailbox.item.saveAsync((result) => {
//       if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
//       else reject(new Error(`Item save karne mein naakam: ${result.error.message}`));
//     });
//   });
// };


// // --- HELPER FUNCTIONS ---

// /**
//  * **DEBUGGING KE LIYE:** Yeh function PDF blob ko browser mein download karta hai.
//  */
// const triggerPdfDownload = (pdfBlob: Blob, filename: string) => {
//   const link = document.createElement('a');
//   const url = URL.createObjectURL(pdfBlob);
//   link.href = url;
//   link.download = filename;
//   document.body.appendChild(link);
//   link.click();
//   document.body.removeChild(link);
//   URL.revokeObjectURL(url);
//   log(`Download shuru kiya gaya: ${filename}`);
// };

// /**
//  * **ZAROORI FIX:** Outlook for Windows client API ke saath aek mehfooz filename banata hai.
//  * Yeh function '()' jaise masla karne wale characters ko hata deta hai.
//  */
// const sanitizeFilename = (name: string): string => {
//   const extension = ".pdf";

//   // baseName without extension
//   let baseName = name.toLowerCase().endsWith(extension)
//     ? name.slice(0, -extension.length)
//     : name.includes(".") ? name.slice(0, name.lastIndexOf(".")) : name;

//   // replace invalid chars
//   let sanitized = baseName.replace(/[\s()\[\]{}]+/g, "_");
//   sanitized = sanitized.replace(/[^a-zA-Z0-9._-]/g, "");

//   // fallback if empty
//   if (!sanitized || sanitized === "_" || sanitized === ".") {
//     sanitized = `converted_${Date.now()}`;
//   }

//   // safe length (200 max instead of 250)
//   if (sanitized.length > 200) {
//     sanitized = sanitized.substring(0, 200);
//   }

//   return sanitized + extension;
// };



// const blobToBase64 = (blob: Blob): Promise<string> => {
//   return new Promise((resolve, reject) => {
//     const reader = new FileReader();
//     reader.onload = () => {
//       const dataUrl = reader.result as string;
//       resolve(dataUrl.split(',')[1]);
//     };
//     reader.onerror = (error) => reject(error);
//     reader.readAsDataURL(blob);
//   });
// };

// // --- MAIN COMPONENT ---
// export const App = () => {
//  const [attachments, setAttachments] = React.useState<IAttachment[]>([]);
//   const [globalStatus, setGlobalStatus] = React.useState<IGlobalStatus>({ message: 'Attachments load ho rahe hain...', type: 'info' });
//   const [isConverting, setIsConverting] = React.useState(false);
//   const [removeOriginal, { toggle: toggleRemoveOriginal }] = useBoolean(true);

//   // --- Data Fetching Effect ---
//   React.useEffect(() => {
//     Office.onReady(async () => {
//       if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
//         await loadAttachments();
//       } else {
//         setGlobalStatus({ message: 'Yeh add-in sirf email items par kaam karta hai.', type: 'warning' });
//       }
//     });
//   }, []);

//   // --- CORE LOGIC ---
//   const loadAttachments = async () => {
//     try {
//       const fetchedAttachments = await getAttachmentsAsync();
//       const compatibleAttachments:any = fetchedAttachments
//         .filter(att => !att.isInline && (att.name.toLowerCase().endsWith('.docx') || att.name.toLowerCase().endsWith('.xlsx')))
//         .map(att => ({
//             id: att.id,
//             name: att.name,
//             newName: sanitizeFilename(att.name),
//             isSelected: true,
//             contentType: att.name.toLowerCase().endsWith('.docx')
//               ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
//               : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
//             status: 'pending',
//             statusMessage: '',
//           }));
        
//       setAttachments(compatibleAttachments);
//       setGlobalStatus({ 
//           message: compatibleAttachments.length > 0 ? `${compatibleAttachments.length} compatible files mili hain.` : 'Koi .docx ya .xlsx attachment nahi mila.', 
//           type: 'info' 
//       });
//     } catch (error) {
//       setGlobalStatus({ message: `Attachments load karne mein error: ${error.message}`, type: 'error' });
//     }
//   };
  
//   const handleConvertClick = async () => {
//     const selectedAttachments = attachments.filter(a => a.isSelected);
//     if (selectedAttachments.length === 0) {
//       setGlobalStatus({ message: 'Convert karne ke liye kam se kam aek file select karein.', type: 'error' });
//       return;
//     }

//     setIsConverting(true);
//     setGlobalStatus({ message: 'Conversion process shuru ho raha hai...', type: 'info' });

//     const results = await Promise.allSettled(selectedAttachments.map(processAttachment));

//     const successCount = results.filter(r => r.status === 'fulfilled').length;
//     const failedCount = results.length - successCount;

//     // --- Aakhri Status Update ---
//     if (failedCount === 0) {
//       setGlobalStatus({ message: `${successCount} file(s) kamyabi se convert ho gayi hain.`, type: 'success' });
//     } else if (successCount > 0) {
//       setGlobalStatus({ message: `Process mukammal: ${successCount} kamyab, ${failedCount} naakam.`, type: 'warning' });
//     } else {
//       setGlobalStatus({ message: 'Sabhi conversions naakam ho gaye. Neeche diye gaye errors check karein.', type: 'error' });
//     }

//     setIsConverting(false);
//     await loadAttachments(); // Sirf baaqi attachments dikhane ke liye list ko refresh karein
//   };

//   const processAttachment = async (attachment: IAttachment) => {
//       let tempItemId: string | null = null;
//       try {
//           updateAttachmentStatus(attachment.id, 'processing', 'Validating...');
//           validateNewFilename(attachment.newName);
//           const token = await Get_Token_SSO();

//           updateAttachmentStatus(attachment.id, 'processing', 'File parh rahe hain...');
//           const fileContent = await getAttachmentContentAsync(attachment.id);
//           validateFileContent(fileContent, attachment.name);
          
//           updateAttachmentStatus(attachment.id, 'processing', 'Cloud par upload ho raha hai...');
//           const uniqueFileNameForUpload = `${Date.now()}_${attachment.name}`;
//           const uploadUrl = `${GRAPH_API_ENDPOINT}/me/drive/root:/${TEMP_ONEDRIVE_FOLDER}/${uniqueFileNameForUpload}:/content`;
//           const uploadResponse = await fetch(uploadUrl, {
//               method: 'PUT',
//               headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': attachment.contentType },
//               body: fileContent
//           });
//           if (!uploadResponse.ok) throw new Error(`OneDrive Upload naakam: ${await uploadResponse.text()}`);
//           const uploadedFile = await uploadResponse.json();
//           tempItemId = uploadedFile.id;

//           updateAttachmentStatus(attachment.id, 'processing', 'PDF mein convert ho raha hai...');
//           const pdfBlob = await convertFileToPdf(token, tempItemId);

//           // <<< --- TESTING KE LIYE: YEH LINE PDF DOWNLOAD KAREGI --- >>>
//           // Testing ke baad is line ko comment kar dein ya hata dein.
//           // triggerPdfDownload(pdfBlob, attachment.newName);
//           // <<< ---------------------------------------------------- >>>

//           if (removeOriginal) {
//               updateAttachmentStatus(attachment.id, 'processing', 'Asal file hata rahe hain...');
//               await removeAttachmentAsync(attachment.id);
//               log('Asal attachment hata di. State sync ke liye item save kar rahe hain...');
//               await saveItemAsync();
//               log('Item save ho gaya. Ab naya attachment shamil kar rahe hain.');
//           }

//           updateAttachmentStatus(attachment.id, 'processing', 'PDF attach ho raha hai...');
//           const currentAttachmentNames = (await getAttachmentsAsync()).map(a => a.name);
//           const uniquePdfName = generateUniqueAttachmentName(attachment.newName, currentAttachmentNames);
//           log(`Mehfooz naam se PDF attach kar rahe hain: ${uniquePdfName}`);
//           const pdfBase64 = await blobToBase64(pdfBlob);
//          await addFileAttachmentFromBase64Async(pdfBase64, uniquePdfName);

          
//           updateAttachmentStatus(attachment.id, 'success', 'Kamyabi se convert ho gaya!');

//       } catch(error) {
//         updateAttachmentStatus(attachment.id, 'error', error.message);
//         throw error; // Promise.allSettled mein pakarne ke liye error ko dobara throw karein
//       } finally {
//           // Cleanup
//           if (tempItemId) {
//               const token = await Get_Token_SSO();
//               const deleteUrl = `${GRAPH_API_ENDPOINT}/me/drive/items/${tempItemId}`;
//               await fetch(deleteUrl, { method: 'DELETE', headers: { 'Authorization': `Bearer ${token}` } });
//               log('Temporary file OneDrive se delete kar di.', { itemId: tempItemId });
//           }
//       }
//   };

  // const convertFileToPdf = async (token: string, itemId: string): Promise<Blob> => {
  //     const convertUrl = `${GRAPH_API_ENDPOINT}/me/drive/items/${itemId}/content?format=pdf`;
  //     const response = await fetch(convertUrl, {
  //         headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/pdf' }
  //     });
  //     if (!response.ok) {
  //         const correlationId = response.headers.get('request-id') || 'N/A';
  //         throw new Error(`PDF Conversion API naakam (Correlation-ID: ${correlationId})`);
  //     }
  //     return await response.blob();
  // };

//   // --- UI HANDLERS & HELPERS ---
  // const updateAttachmentStatus = (id: string, status: AttachmentStatus, message: string) => {
  //   setAttachments(prev => prev.map(a => a.id === id ? { ...a, status, statusMessage: message } : a));
  // };
  
  // const handleNameChange = (id: string, newName: string) => {
  //   setAttachments(prev => prev.map(att => att.id === id ? { ...att, newName: newName } : att));
  // };
  
  // const handleToggleSelect = (id: string) => {
  //   setAttachments(prev => prev.map(a => a.id === id ? { ...a, isSelected: !a.isSelected } : a));
  // };

  // const validateNewFilename = (value: string) => {
  //   if (!value || value.trim() === '') throw new Error('Filename khaali nahi ho sakta.');
  //   if (!value.toLowerCase().endsWith('.pdf')) throw new Error('Filename .pdf par khatam hona chahiye.');
  //   if (value.length > MAX_FILENAME_LENGTH) throw new Error(`Filename bohot lamba hai (max ${MAX_FILENAME_LENGTH} chars).`);
  // };
  
  // const getFilenameErrorMessage = (value: string): string => {
  //   try {
  //     validateNewFilename(value);
  //     const sanitized = sanitizeFilename(value);
  //     if(sanitized !== value) {
  //       return `Is tarah save hoga: ${sanitized}`;
  //     }
  //     return '';
  //   } catch (error) {
  //     return error.message;
  //   }
  // };

  // const validateFileContent = (content: ArrayBuffer, filename: string): void => {
  //   if (!content || content.byteLength === 0) throw new Error(`File "${filename}" khaali hai.`);
  //   if (content.byteLength > MAX_FILE_SIZE_MB * 1024 * 1024) throw new Error(`File ${MAX_FILE_SIZE_MB}MB se bari hai.`);
  //   const view = new DataView(content);
  //   if (view.getUint32(0, true) !== 0x04034b50) throw new Error(`File aek valid Office OpenXML file nahi hai.`);
  // };

  // const generateUniqueAttachmentName = (desiredName: string, existingNames: string[]): string => {
  //   let uniqueName = sanitizeFilename(desiredName);
  //   if (!existingNames.includes(uniqueName)) return uniqueName;
    
  //   let counter = 1;
  //   const nameWithoutExt = uniqueName.substring(0, uniqueName.lastIndexOf('.'));
  //   const extension = uniqueName.substring(uniqueName.lastIndexOf('.'));
  //   while (existingNames.includes(uniqueName)) {
  //     uniqueName = `${nameWithoutExt}_${counter}${extension}`;
  //     counter++;
  //   }
  //   return uniqueName;
  // };
  

//   // --- RENDER ---
//   return (
//     <Stack tokens={{ childrenGap: 15 }} style={{ padding: 20 }}>
//       <Text variant="xLarge" styles={{ root: { fontWeight: '600' } }}>Attachment PDF Converter</Text>

//       {globalStatus.message && <MessageBar messageBarType={MessageBarType[globalStatus.type]}>{globalStatus.message}</MessageBar>}

//       <Stack tokens={{ childrenGap: 10 }}>
//         {attachments.length === 0 && !isConverting && <Text>Koi compatible attachments (.docx, .xlsx) nahi mili hain.</Text>}
        
//         {attachments.map(att => (
//           <Stack key={att.id} tokens={{ childrenGap: 5 }} styles={{ root: { border: '1px solid #eee', padding: 10, borderRadius: 4, opacity: isConverting && !att.isSelected ? 0.5 : 1 } }}>
//             <Checkbox label={att.name} checked={att.isSelected} onChange={() => handleToggleSelect(att.id)} disabled={isConverting} />
            
//             <TextField
//               label="Naya PDF Filename:"
//               value={att.newName}
//               onChange={(_e, val) => handleNameChange(att.id, val || '')}
//               disabled={isConverting || !att.isSelected}
//               onGetErrorMessage={getFilenameErrorMessage}
//               validateOnFocusOut
//               description="Ghair zaroori characters convert karte waqt khud hi hata diye jayenge."
//             />

//             {att.status !== 'pending' && (
//               <MessageBar messageBarType={MessageBarType[att.status]}>
//                   {att.status === 'processing' && <Spinner size={SpinnerSize.xSmall} style={{ marginRight: 8 }} />}
//                   {att.statusMessage}
//               </MessageBar>
//             )}
//           </Stack>
//         ))}
//       </Stack>

//       {attachments.length > 0 && (
//         <Stack tokens={{ childrenGap: 10 }}>
//           <Checkbox label="Conversion ke baad asal files hata dein" checked={removeOriginal} onChange={toggleRemoveOriginal} disabled={isConverting} />
          
//           <PrimaryButton 
//             onClick={handleConvertClick} 
//             disabled={isConverting || attachments.filter(a => a.isSelected).length === 0}
//           >
//             {isConverting && <Spinner size={SpinnerSize.small} style={{ marginRight: 8 }} />}
//             {isConverting ? 'Convert ho raha hai...' : `${attachments.filter(a => a.isSelected).length} Selected Files ko PDF banayein`}
//           </PrimaryButton>
//         </Stack>
//       )}
//     </Stack>
//   );
// };










import * as React from 'react';
import {
  Text,
  PrimaryButton,
  Checkbox,
  TextField,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  initializeIcons,
  Icon,
  mergeStyleSets,
  FontWeights,
  ProgressIndicator, // New import for progress bar
  IconButton,       // New import for edit button
  Callout,          // New import for showing errors
} from '@fluentui/react';
import { useBoolean, useId } from '@fluentui/react-hooks';
import { Get_Token_SSO } from './SSO_For_Graph';

// Initialize Fluent UI icons
initializeIcons();

// --- CONSTANTS ---
const MAX_FILENAME_LENGTH = 250;
const MAX_FILE_SIZE_MB = 10;
const GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0";
const TEMP_ONEDRIVE_FOLDER = "Temp-add-in-uploads";

// --- TYPES & INTERFACES (UNCHANGED) ---
type AttachmentStatus = 'pending' | 'processing' | 'success' | 'error';
type GlobalStatusType = 'info' | 'success' | 'error' | 'warning';

interface IAttachment {
  id: string;
  name: string;
  size: number;
  newName: string;
  isSelected: boolean;
  contentType: string;
  status: AttachmentStatus;
  statusMessage: string;
}

interface IGlobalStatus {
  message: string;
  type: GlobalStatusType;
}

// --- LOGGING & OFFICE.JS WRAPPERS (UNCHANGED) ---
// ... (All your original helper functions: log, getAttachmentsAsync, etc. remain here)
// --- Omitted for brevity, paste your original functions here ---
const log = (message: string, data?: any) => { console.log(`[PDF Add-in] ${message}`, data !== undefined ? data : ''); };
const getAttachmentsAsync = (): Promise<Office.AttachmentDetails[]> => { return new Promise((resolve, reject) => { Office.context.mailbox.item.getAttachmentsAsync((result: any) => { if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value); else reject(new Error(result.error.message)); }); }); };
const getAttachmentContentAsync = (id: string): Promise<ArrayBuffer> => { return new Promise((resolve, reject) => { Office.context.mailbox.item.getAttachmentContentAsync(id, (result) => { if (result.status === Office.AsyncResultStatus.Succeeded) { try { const binaryString = window.atob(result.value.content); const bytes = new Uint8Array(binaryString.length); for (let i = 0; i < binaryString.length; i++) { bytes[i] = binaryString.charCodeAt(i); } resolve(bytes.buffer); } catch (error) { reject(new Error(`Base64 content decode karne mein naakam: ${error.message}`)); } } else { reject(new Error(result.error.message)); } }); }); };
const addFileAttachmentFromBase64Async = (base64: string, name: string): Promise<string> => { return new Promise((resolve, reject) => { Office.context.mailbox.item.addFileAttachmentFromBase64Async( base64, name, { isInline: false }, (asyncResult) => { if (asyncResult.status === Office.AsyncResultStatus.Succeeded) { resolve(asyncResult.value); } else { reject(new Error(`Attachment failed. Code: ${asyncResult.error.code}, Message: ${asyncResult.error.message}`)); } } ); }); };
const removeAttachmentAsync = (id: string): Promise<void> => { return new Promise((resolve) => { Office.context.mailbox.item.removeAttachmentAsync(id, (result) => { if (result.status !== Office.AsyncResultStatus.Succeeded) { console.warn(`Original attachment not removed, but process continues. Error: ${result.error.message}`); } resolve(); }); }); };
const saveItemAsync = (): Promise<void> => { return new Promise((resolve, reject) => { Office.context.mailbox.item.saveAsync((result) => { if (result.status === Office.AsyncResultStatus.Succeeded) resolve(); else reject(new Error(`Failed to save item: ${result.error.message}`)); }); }); };
const sanitizeFilename = (name: string): string => { const extension = ".pdf"; let baseName = name.toLowerCase().endsWith(extension) ? name.slice(0, -extension.length) : name.includes(".") ? name.slice(0, name.lastIndexOf(".")) : name; let sanitized = baseName.replace(/[\s()\[\]{}]+/g, "_").replace(/[^a-zA-Z0-9._-]/g, ""); if (!sanitized || sanitized === "_" || sanitized === ".") sanitized = `converted_${Date.now()}`; if (sanitized.length > 200) sanitized = sanitized.substring(0, 200); return sanitized + extension; };
const blobToBase64 = (blob: Blob): Promise<string> => { return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = () => { const dataUrl = reader.result as string; resolve(dataUrl.split(',')[1]); }; reader.onerror = (error) => reject(error); reader.readAsDataURL(blob); }); };


// --- NEW STYLING: MODERN, UNIQUE, AND VIBRANT ---
const classNames = mergeStyleSets({
  taskpane: {
    display: 'flex',
    flexDirection: 'column',
    height: '100vh',
    backgroundColor: '#f7f9fc',
  },
  header: {
    padding: '12px 20px',
    display: 'flex',
    alignItems: 'center',
    background: 'linear-gradient(45deg, #ffffff 0%, #f9fbff 100%)',
    borderBottom: '1px solid #e1e5ea',
    flexShrink: 0,
  },
  headerLogo: {
    width: 36, height: 36,
    background: 'linear-gradient(135deg, #489dffff 0%, #4891ff 100%)',
    borderRadius: 10,
    marginRight: 12,
    display: 'flex', alignItems: 'center', justifyContent: 'center',
    boxShadow: '0 4px 10px rgba(110, 72, 255, 0.2)',
  },
  headerTitle: {
    fontWeight: FontWeights.semibold,
    fontSize: 20,
    color: '#1a202c',
  },
  content: {
    flex: 1,
    overflowY: 'auto',
    padding: '10px',
    '& > *:not(:last-child)': {
      marginBottom: '12px',
    },
  },
  attachmentItem: {
    backgroundColor: '#ffffff',
    borderRadius: 12,
    border: '1px solid #e1e5ea',
    padding: '12px',
    display: 'grid',
    gridTemplateColumns: 'auto 48px 1fr auto',
    alignItems: 'center',
    // gap: 12,
    transition: 'all 0.2s cubic-bezier(0.25, 0.8, 0.25, 1)',
    selectors: {
      '&:hover': {
        transform: 'translateY(-2px)',
        boxShadow: '0 8px 25px rgba(0, 0, 0, 0.07)',
        borderColor: '#c9d3e0',
      },
      '& .editButton': {
        opacity: 0,
        transition: 'opacity 0.2s ease',
      },
      '&:hover .editButton': {
        opacity: 1,
      },
    },
  },
  itemSelected: {
    borderColor: '#169feeff',
    // boxShadow: '0 0 0 2px rgba(110, 72, 255, 0.2)',
  },
  fileIcon: {
    width: 48, height: 48,
    borderRadius: 10,
    display: 'flex', alignItems: 'center', justifyContent: 'center',
    fontSize: 24,
  },
  wordBg: { background: 'rgba(43, 87, 154, 0.1)', color: '#2B579A' },
  excelBg: { background: 'rgba(33, 115, 70, 0.1)', color: '#217346' },
  fileDetails: {
    display: 'flex', flexDirection: 'column',
    minWidth: 0, // Prevents text overflow issues
  },
  fileNameContainer: {
    display: 'flex', alignItems: 'center', gap: '4px'
  },
  fileName: {
    fontWeight: FontWeights.semibold, color: '#1a202c',
    whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
  },
  fileMeta: {
    fontSize: 12, color: '#718096',
  },
  statusIndicator: {
    display: 'flex', alignItems: 'center',
    gap: 6, padding: '4px 10px', borderRadius: 16,
    fontSize: 12, fontWeight: FontWeights.semibold,
  },
  statusSuccess: { background: 'rgba(16, 124, 16, 0.1)', color: '#107C10',padding:'5px' },
  statusError: { background: 'rgba(217, 21, 28, 0.1)', color: '#d9151c', cursor: 'pointer' },
  statusProcessing: { background: 'rgba(0, 90, 158, 0.1)', color: '#005A9E',padding:'5px' },
  
  actionBar: {
    padding: '16px 20px',
    borderTop: '1px solid #e1e5ea',
    backgroundColor: 'rgba(255, 255, 255, 0.8)',
    backdropFilter: 'blur(10px)',
    boxShadow: '0 -5px 20px rgba(0, 0, 0, 0.05)',
    flexShrink: 0,
  },
  actionButton: {
    width: '100%',
    height: 48,
    borderRadius: 10,
    fontWeight: FontWeights.semibold, fontSize: 16,
    background: 'linear-gradient(135deg, #6e48ff 0%, #4891ff 100%)',
    border: 'none',
    boxShadow: '0 4px 14px rgba(110, 72, 255, 0.3)',
    selectors: {
      '&:hover': {
        background: 'linear-gradient(135deg, #613de0 0%, #3a82e8 100%)',
        boxShadow: '0 6px 16px rgba(110, 72, 255, 0.4)',
      },
    },
  },
  options: {
    marginBottom: 16,
  },
  emptyState: {
    textAlign: 'center', color: '#a0aec0',
    padding: '60px 20px',
  },
  emptyStateIcon: {
    fontSize: 64, color: '#e2e8f0',
  },
});

// --- MAIN COMPONENT ---
export const App = () => {
  const [attachments, setAttachments] = React.useState<IAttachment[]>([]);
  const [globalStatus, setGlobalStatus] = React.useState<IGlobalStatus | null>({ message: 'Searching for compatible attachments...', type: 'info' });
  const [isConverting, setIsConverting] = React.useState(false);
  const [removeOriginal, { toggle: toggleRemoveOriginal }] = useBoolean(true);
  
  // NEW state for a more interactive UI
  const [editingId, setEditingId] = React.useState<string | null>(null);
  const [errorCallout, setErrorCallout] = React.useState<{ target: HTMLElement, message: string } | null>(null);
  const [progress, setProgress] = React.useState(0);

  // --- Data Fetching Effect (Unchanged) ---
   React.useEffect(() => {
    Office.onReady(async () => {
      if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
        await loadAttachments();
      } else {
        setGlobalStatus({ message: 'This add-in only works on email items.', type: 'warning' });
      }
    });
  }, []);
  // --- CORE LOGIC (With Progress Updates) ---
   const loadAttachments = async () => {
    try {

      // List of all supported extensions
const SUPPORTED_EXTENSIONS = [
  '.docx', '.doc', '.dot', '.dotx', 
  '.xlsx', '.xls', '.xlsb', '.xlsm',
  '.pptx', '.ppt', '.pot', '.potx', '.pps', '.ppsx',
  '.odt', '.ods', '.odp',
  '.csv', '.rtf'
];

const MIME_TYPE_MAP = {
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.doc': 'application/msword',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.xls': 'application/vnd.ms-excel',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.ppt': 'application/vnd.ms-powerpoint',
  '.csv': 'text/csv',
  '.rtf': 'application/rtf',
  '.odt': 'application/vnd.oasis.opendocument.text',
  '.ods': 'application/vnd.oasis.opendocument.spreadsheet',
  '.odp': 'application/vnd.oasis.opendocument.presentation',
  // Add other supported types here
};


const getContentTypeFromFileName = (fileName: string): string => {
    const extension = '.' + fileName.split('.').pop()?.toLowerCase();
    return MIME_TYPE_MAP[extension] || 'application/octet-stream'; // Default for unknown types
};


      const fetchedAttachments = await getAttachmentsAsync();
      const compatibleAttachments: IAttachment[] = fetchedAttachments
        .filter(att => {
    const lowerCaseName = att.name.toLowerCase();
    return !att.isInline && SUPPORTED_EXTENSIONS.some(ext => lowerCaseName.endsWith(ext));
})
        .map(att => ({
            id: att.id,
            name: att.name,
            size: att.size,
            newName: sanitizeFilename(att.name),
            isSelected: true,
            contentType: getContentTypeFromFileName(att.name),
            status: 'pending',
            statusMessage: '',
          }));
        
      setAttachments(compatibleAttachments);
      setGlobalStatus({ 
          message: compatibleAttachments.length > 0 ? `${compatibleAttachments.length} compatible file(s) found.` : 'No compatible .docx or .xlsx attachments found.', 
          type: compatibleAttachments.length > 0 ? 'info' : 'warning'
      });
    } catch (error) {
      setGlobalStatus({ message: `Error loading attachments: ${error.message}`, type: 'error' });
    }
  };
  

  const handleConvertClick = async () => {
    const selectedAttachments = attachments.filter(a => a.isSelected);
    if (selectedAttachments.length === 0) {
      setGlobalStatus({ message: 'Please select at least one file to convert.', type: 'warning' });
      return;
    }

    setIsConverting(true);
    setProgress(0);
    setGlobalStatus({ message: 'Starting conversion...', type: 'info' });

    // NEW: Calculate total steps for the progress bar
    const totalSteps = selectedAttachments.length * 5; // 5 steps per file
    let completedSteps = 0;

    const updateProgress = () => {
      completedSteps++;
      setProgress((completedSteps / totalSteps) * 100);
    };

    const results = await Promise.allSettled(
      selectedAttachments.map(att => processAttachment(att, updateProgress))
    );

    // ... same result handling logic ...
  };

  const processAttachment = async (attachment: IAttachment, onProgress: () => void) => {
    let tempItemId: string | null = null;
    try {
      updateAttachmentStatus(attachment.id, 'processing', 'Validating...');
      validateNewFilename(attachment.newName);
      const token = await Get_Token_SSO();
      console.log(token);
      
      onProgress(); // Step 1

      updateAttachmentStatus(attachment.id, 'processing', 'Uploading to cloud...');
      const fileContent = await getAttachmentContentAsync(attachment.id);
      validateFileContent(fileContent, attachment.name);
      const uniqueFileNameForUpload = `${Date.now()}_${attachment.name}`;
      const uploadUrl = `${GRAPH_API_ENDPOINT}/me/drive/root:/${TEMP_ONEDRIVE_FOLDER}/${uniqueFileNameForUpload}:/content`;
        const uploadResponse = await fetch(uploadUrl, {
              method: 'PUT',
              headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': attachment.contentType },
              body: fileContent
          });
      if (!uploadResponse.ok) throw new Error(`OneDrive Upload failed: ${await uploadResponse.text()}`);
      const uploadedFile = await uploadResponse.json();
      tempItemId = uploadedFile.id;
      onProgress(); // Step 2

      updateAttachmentStatus(attachment.id, 'processing', 'Converting to PDF...');
      const Newtoken = await Get_Token_SSO();
      const pdfBlob = await convertFileToPdf(Newtoken, tempItemId);
      onProgress(); // Step 3

      if (removeOriginal) {
        updateAttachmentStatus(attachment.id, 'processing', 'Tidying up...');
        await removeAttachmentAsync(attachment.id);
      }
      await saveItemAsync();
      onProgress(); // Step 4
      
      updateAttachmentStatus(attachment.id, 'processing', 'Attaching PDF...');
      const currentAttachmentNames = (await getAttachmentsAsync()).map(a => a.name);
      const uniquePdfName = generateUniqueAttachmentName(attachment.newName, currentAttachmentNames);
      const pdfBase64 = await blobToBase64(pdfBlob);
      await addFileAttachmentFromBase64Async(pdfBase64, uniquePdfName);
      onProgress(); // Step 5

      updateAttachmentStatus(attachment.id, 'success', 'Converted!');
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      updateAttachmentStatus(attachment.id, 'error', errorMessage);
      throw error;
    } finally {
      if (tempItemId) { /* ... same cleanup logic ... */ }
    }
  };

  // --- Other helper functions (unchanged) ---

  const convertFileToPdf = async (token: string, itemId: string): Promise<Blob> => {
    console.log("token in convertFileToPdf:", token);
    
      const convertUrl = `${GRAPH_API_ENDPOINT}/me/drive/items/${itemId}/content?format=pdf`;
      const response = await fetch(convertUrl, {
          headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/pdf' }
      });
      if (!response.ok) {
          const correlationId = response.headers.get('request-id') || 'N/A';
          throw new Error(`PDF Conversion API failed (Correlation-ID: ${correlationId})`);
      }
      return await response.blob();
  };
    const updateAttachmentStatus = (id: string, status: AttachmentStatus, message: string) => {
    setAttachments(prev => prev.map(a => a.id === id ? { ...a, status, statusMessage: message } : a));
  };
  const handleToggleSelect = (id: string) => {
    setAttachments(prev => prev.map(a => a.id === id ? { ...a, isSelected: !a.isSelected } : a));
  };
  const validateNewFilename = (value: string) => {
    if (!value || value.trim() === '') throw new Error('Filename cannot be empty.');
    if (!value.toLowerCase().endsWith('.pdf')) throw new Error('Filename must end with .pdf.');
    if (value.length > MAX_FILENAME_LENGTH) throw new Error(`Filename is too long (max ${MAX_FILENAME_LENGTH} chars).`);
  };
    const validateFileContent = (content: ArrayBuffer, filename: string): void => {
    if (!content || content.byteLength === 0) throw new Error(`File "${filename}" is empty.`);
    if (content.byteLength > MAX_FILE_SIZE_MB * 1024 * 1024) throw new Error(`File exceeds the ${MAX_FILE_SIZE_MB}MB size limit.`);
    const view = new DataView(content);
    if (view.getUint32(0, false) !== 0x504B0304) log(`File "${filename}" might not be a valid Office OpenXML file (ZIP header not found).`);
  };

  const generateUniqueAttachmentName = (desiredName: string, existingNames: string[]): string => {
    let uniqueName = sanitizeFilename(desiredName);
    if (!existingNames.find(name => name.toLowerCase() === uniqueName.toLowerCase())) return uniqueName;
    
    let counter = 1;
    const nameWithoutExt = uniqueName.substring(0, uniqueName.lastIndexOf('.'));
    const extension = uniqueName.substring(uniqueName.lastIndexOf('.'));
    while (existingNames.find(name => name.toLowerCase() === uniqueName.toLowerCase())) {
      uniqueName = `${nameWithoutExt}_${counter}${extension}`;
      counter++;
    }
    return uniqueName;
  };



  


  // NEW: UI handler for inline editing
  const handleNameChange = (id: string, newName: string) => {
    setAttachments(prev => prev.map(att => att.id === id ? { ...att, newName } : att));
  };
  const handleEditBlur = (id: string) => {
    const attachment = attachments.find(att => att.id === id);
    if(attachment) {
      try {
        validateNewFilename(attachment.newName);
      } catch (e) {
        // If invalid on blur, revert to a valid name
        handleNameChange(id, sanitizeFilename(attachment.name));
      }
    }
    setEditingId(null);
  };
  
  const selectedCount = attachments.filter(a => a.isSelected).length;


// Example classNames object
const IconclassNames = {
  wordBg: 'bg-blue-100 text-blue-700',       // Word
  excelBg: 'bg-green-100 text-green-700',    // Excel
  pptBg: 'bg-orange-100 text-orange-700',    // PowerPoint
  genericBg: 'bg-gray-100 text-gray-700',    // Generic
};

const getFileIconProps = (fileName: string) => {
  const lowerCaseName = fileName.toLowerCase();

  if (lowerCaseName.includes('.doc'))
    return { icon: 'WordDocument', bgClass: IconclassNames.wordBg };

  if (lowerCaseName.includes('.xls'))
    return { icon: 'ExcelDocument', bgClass: IconclassNames.excelBg };

  if (lowerCaseName.includes('.ppt'))
    return { icon: 'PowerPointDocument', bgClass: IconclassNames.pptBg };

  // Default icon for all other supported types
  return { icon: 'Page', bgClass: IconclassNames.genericBg };
};








  // --- RENDER ---
  return (
    <div className={classNames.taskpane}>
      {/* HEADER */}
      <header className={classNames.header}>
        <div className={classNames.headerLogo}>
          <Icon iconName="PDF" styles={{ root: { color: '#ffffff', fontSize: 20 } }} />
        </div>
        <Text className={classNames.headerTitle}>PDF Power Converter</Text>
      </header>
      
      {/* GLOBAL STATUS (subtler placement) */}
      {globalStatus && !isConverting && (
        <div style={{ padding: '0 20px', marginTop: '16px' }}>
          <MessageBar messageBarType={MessageBarType[globalStatus.type]} onDismiss={() => setGlobalStatus(null)} isMultiline={false}>
            {globalStatus.message}
          </MessageBar>
        </div>
      )}

      {/* CONTENT: ATTACHMENT LIST */}
       <main className={classNames.content}>
        {attachments.length === 0 && !isConverting ? (
          <div className={classNames.emptyState}>
            <Icon iconName="OpenFile" className={classNames.emptyStateIcon} />
            <Text variant="large" styles={{ root: { fontWeight: FontWeights.semibold, color: '#4a5568' } }}>Ready for Action</Text>
            <Text styles={{ root: { marginTop: 8 } }}>No convertible files were found in this email.</Text>
          </div>
        ) : attachments.map(att => {
            // --- UPDATED: Use the dynamic icon helper ---
            const iconProps = getFileIconProps(att.name);
            return (
              <div key={att.id} className={`${classNames.attachmentItem} ${att.isSelected ? classNames.itemSelected : ''}`}>
                <Checkbox checked={att.isSelected} onChange={() => handleToggleSelect(att.id)} disabled={isConverting} />
                
                <div className={`${classNames.fileIcon} ${iconProps.bgClass}`}>
                  <Icon iconName={iconProps.icon} />
                </div>
                
                <div className={classNames.fileDetails}>
                  {editingId === att.id ? (
                    <TextField value={att.newName} onChange={(_e, val) => handleNameChange(att.id, val || '')} onBlur={() => handleEditBlur(att.id)} autoFocus styles={{ fieldGroup: { height: 28, borderRadius: 6 } }} />
                  ) : (
                    <div className={classNames.fileNameContainer}>
                      <Text block className={classNames.fileName} title={att.name}>{att.newName}</Text>
                      {!isConverting && (
                        <IconButton className="editButton" iconProps={{ iconName: 'Edit' }} title="Rename" ariaLabel="Rename" onClick={() => setEditingId(att.id)} styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 12 } }} />
                      )}
                    </div>
                  )}
                  <Text block className={classNames.fileMeta}>{(att.size / 1024).toFixed(1)} KB &middot; {att.name.split('.').pop()}</Text>
                </div>
                
                <div className={classNames.statusIndicator}>
                  {att.status === 'processing' && <><Spinner size={SpinnerSize.xSmall} /><span className={classNames.statusProcessing}>{att.statusMessage}</span></>}
                  {att.status === 'success' && <><Icon iconName="CheckMark" /><span className={classNames.statusSuccess}>Done</span></>}
                  {att.status === 'error' && (
                    <div id={`error-target-${att.id}`} className={classNames.statusError} onClick={(e) => setErrorCallout({ target: e.currentTarget, message: att.statusMessage })}>
                      <Icon iconName="Warning" /><span  className={classNames.statusSuccess}>Error</span>
                    </div>
                  )}
                </div>
              </div>
            )
          })
        }
      </main>


      {/* ACTION BAR: STICKY FOOTER */}
      {attachments.length > 0 && (
        <footer className={classNames.actionBar}>
          {isConverting ? (
            <ProgressIndicator 
              label={`Converting ${selectedCount} file(s)...`} 
              description={`${Math.round(progress)}% Complete`} 
              percentComplete={progress / 100}
            />
          ) : (
            <>
              <div className={classNames.options}>
                <Checkbox label="Remove original files" checked={removeOriginal} onChange={toggleRemoveOriginal} />
              </div>
              <PrimaryButton 
                onClick={handleConvertClick} 
                disabled={selectedCount === 0}
                text={`Convert ${selectedCount} Selected File(s)`}
                className={classNames.actionButton}
              />
            </>
          )}
        </footer>
      )}
      
      {/* Error Callout for details on demand */}
      {errorCallout && (
          <Callout
              target={errorCallout.target}
              onDismiss={() => setErrorCallout(null)}
              role="alert"
              setInitialFocus
          >
              <div style={{ padding: '12px 20px', maxWidth: 300 }}>
                <Text block variant="mediumPlus" styles={{root: {fontWeight: FontWeights.semibold, marginBottom: 8}}}>Conversion Error</Text>
                <Text>{errorCallout.message}</Text>
              </div>
          </Callout>
      )}
    </div>
  );
};