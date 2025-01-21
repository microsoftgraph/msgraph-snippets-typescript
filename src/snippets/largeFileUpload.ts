// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Client,
  FileUpload,
  LargeFileUploadTask,
  LargeFileUploadTaskOptions,
  OneDriveLargeFileUploadOptions,
  OneDriveLargeFileUploadTask,
  UploadResult,
} from '@microsoft/microsoft-graph-client';
import { DriveItem, Message } from '@microsoft/microsoft-graph-types';
import { readFile } from 'fs/promises';
import { basename } from 'path';

export default async function runLargeFileUploadSamples(
  graphClient: Client,
  filePath: string,
): Promise<void> {
  const targetFolderPath = 'Documents';

  await uploadFileToOneDrive(graphClient, filePath, targetFolderPath);
  await uploadAttachmentToMessage(graphClient, filePath);
}

async function uploadFileToOneDrive(
  graphClient: Client,
  filePath: string,
  targetFolderPath: string,
): Promise<void> {
  // <LargeFileUploadSnippet>
  // readFile from fs/promises
  const file = await readFile(filePath);
  // basename from path
  const fileName = basename(filePath);

  const options: OneDriveLargeFileUploadOptions = {
    // Relative path from root folder
    path: targetFolderPath,
    fileName: fileName,
    rangeSize: 1024 * 1024,
    uploadEventHandlers: {
      // Called as each "slice" of the file is uploaded
      progress: (range, _) => {
        console.log(`Uploaded bytes ${range?.minValue} to ${range?.maxValue}`);
      },
    },
  };

  // Create FileUpload object
  const fileUpload = new FileUpload(file, fileName, file.byteLength);
  // Create a OneDrive upload task
  const uploadTask = await OneDriveLargeFileUploadTask.createTaskWithFileObject(
    graphClient,
    fileUpload,
    options,
  );

  // Do the upload
  const uploadResult: UploadResult = await uploadTask.upload();

  // The response body will be of the corresponding type of the
  // item being uploaded. For OneDrive, this is a DriveItem
  const driveItem = uploadResult.responseBody as DriveItem;
  console.log(`Uploaded file with ID: ${driveItem.id}`);
  // </LargeFileUploadSnippet>
}

// eslint-disable-next-line no-unused-vars
async function resumeUpload(
  uploadTask: OneDriveLargeFileUploadTask<Blob>,
): Promise<DriveItem> {
  // <ResumeSnippet>
  const resumedFile = (await uploadTask.resume()) as DriveItem;
  // </ResumeSnippet>

  return resumedFile;
}

async function uploadAttachmentToMessage(
  graphClient: Client,
  filePath: string,
): Promise<void> {
  // <UploadAttachmentSnippet>
  // readFile from fs/promises
  const file = await readFile(filePath);
  // basename from path
  const fileName = basename(filePath);

  const options: LargeFileUploadTaskOptions = {
    rangeSize: 1024 * 1024,
    uploadEventHandlers: {
      // Called as each "slice" of the file is uploaded
      progress: (range, _) => {
        console.log(`Uploaded bytes ${range?.minValue} to ${range?.maxValue}`);
      },
    },
  };

  // Create a draft message
  const message: Message = await graphClient.api('/me/messages').post({
    subject: 'Large file attachment',
  });

  // Create upload session using draft message's ID
  const uploadUrl = `/me/messages/${message.id}/attachments/createUploadSession`;
  const uploadSession = await LargeFileUploadTask.createUploadSession(
    graphClient,
    uploadUrl,
    {
      AttachmentItem: {
        attachmentType: 'file',
        name: fileName,
        size: file.byteLength,
      },
    },
  );

  // Create file upload
  const fileUpload = new FileUpload(file, fileName, file.byteLength);

  // Create upload task
  const uploadTask = new LargeFileUploadTask(
    graphClient,
    fileUpload,
    uploadSession,
    options,
  );

  // Upload the file
  const uploadResult = await uploadTask.upload();
  console.log(`File uploaded to ${uploadResult.location}`);
  // </UploadAttachmentSnippet>
}
