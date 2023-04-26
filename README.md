# PSPDFKit---Microsoft-Dynamics-Editor-Sample
The PowerApps Component Framework (PCF) Control retrieves notes containing PDF attachments and presents the chosen PDF utilizing PSPDFKit Web SDK.


## What is PowerApps Component Framework(PCF) ## 
Power Apps Component Framework (PCF) is a framework that allows developers to create custom components for Power Apps and Dynamics 365. With PCF, developers can extend the functionality of these applications by creating custom controls that can be reused across different applications and screens


## Introduction to PSPDFKit Dynamics Editor Component ## 
PSPDFKit provides a client-side JavaScript library that allows developers to integrate PDF viewing, annotation, and editing capabilities into their model-driven applications within Power Apps or Dynamics. By utilizing this library, users can open, edit, and save PDFs directly from a web browser and store the updated PDF data in DataVerse Table.


#### Prerequisites 1 - Hosting the PSPDFKit Binaries in an Azure Storge ####

Steps for creating a storage account and hosting a file in the $web folder in Azure:

1. Sign in to the Azure portal at https://portal.azure.com/
2. Click on "Create a resource" and search for "Storage account".
3. Select the subscription, resource group, and storage account name. Choose the location and performance tier according to your requirements.
4. Under the "Advanced" tab, enable "Hierarchical namespace" and "Blob public access".
5. Click on "Review + create" and then "Create" to create the storage account.
6. Once the storage account is created, select it and go to the "Blobs" section.
7. Create a new "Folder" in the in the $web folder for the storage account.
8. Click on "Upload" to upload the files you want to host. Make sure to set the "Blob type" to "Block blob".
9. Once the files is uploaded, select it and click on "Copy URL" to get the URL of the file.
10. To host the files in the $web folder, go to the "Containers" section and create a new container with the name "$web".
11. Select the $web container and click on "Upload" to upload the file again. This time, set the "Blob type" to "Page blob".
12. Once the files are uploaded to the $web container, copy the URL of the file again and replace the container name with "$web".
13. The files are now hosted in the $web folder and can be accessed using the URL obtained in the previous step.

![image](https://user-images.githubusercontent.com/25176106/234568006-1fd31be7-a47d-450d-a68c-10ce530e2deb.png)


We will include the Storage account link in section below -

```
PSPDFKit.load({
            disableWebAssemblyStreaming: true,
            baseUrl : "<Storage account folder link>",
            licenseKey: this._context.parameters.psPdfKitLicenseKey.raw,
            container: ".pspdfkit-container",
			document: this.arrBuffer,                       
        })
```



#### Prerequisites 2 - PSPDFKit License file ####

Contact PDFPFKit [sales](https://pspdfkit.com/sales/ "sales team")

You can update the license file when you bind the component to a field in the Form.


## Importing and Adding your component on the Form ##









