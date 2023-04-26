import {IInputs, IOutputs} from "./generated/ManifestTypes";

class EntityReference {
	id: string;
	typeName: string;
	constructor(typeName: string, id: string) {
		this.id = id;
		this.typeName = typeName;
	}
}
class AttachedFile implements ComponentFramework.FileObject {
    private _minImageHeight: number | null;
	private _minImageWidth: number | null;
	annotationId: any;
	fileContent: string;
	fileSize: number;
	fileName: string;
	mimeType: string;
	subject:string;
	description:string;
	constructor(annotationId:any,fileName: string, mimeType: string, fileContent: string, fileSize: number, subject:string, description:string) {
		this.annotationId = annotationId;
		this.fileName = fileName;
		this.mimeType = mimeType;
		this.fileContent = fileContent;
		this.fileSize = fileSize;
		this.description = description; 
		this.subject = subject;

	}
}

export class PSPDFKitEditor implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;
	private btn :any;
	private div :HTMLElement;
	public arrBuffer:ArrayBuffer;
	private newCard:HTMLElement;
	private notes:HTMLDivElement;
    private overlayDiv:any;
    private loading:HTMLElement;
    private _notesContainer: HTMLDivElement;

	// PSPDFKit Instance
    private _instance: any;

    /**
     * Empty constructor.
     */
    constructor()
    {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public async init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
    {
        // Add control initialization code		
        this._context = context;
        this._container = container;
  
        let reference: EntityReference = new EntityReference(
            (<any>context).page.entityTypeName,
            (<any>context).page.entityId
        )

        this.loading = document.createElement("div");
        let notesContainer = document.createElement("div");     
        this._notesContainer = notesContainer;

        let notesHeader = document.createElement("div");
        notesHeader.textContent = reference.typeName == "email" ? "Attachments" : "Notes";
        notesContainer.appendChild(notesHeader);

        this.notes = document.createElement("div");
        notesContainer.appendChild(this.notes);
        this._container.appendChild(notesContainer);

        this.div = document.createElement("div");   
		this.div.classList.add("pspdfkit-container");
		this.div.style.height = "100px";
		this.div.style.width = "100px"; 
		this._container.appendChild(this.div);

        let attachedFiles:AttachedFile[] = await this.GetFiles(reference);
		if (attachedFiles){
			this.RenderThumbnails(attachedFiles, this.notes);
		}
    }
    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public async updateView(context: ComponentFramework.Context<IInputs>)
    {
        // Add code to update control view
         // Add code to update control view
		const PSPDFKit = await import("../node_modules/pspdfkit/");

        if (context.updatedProperties[context.updatedProperties.length - 1] == "fullscreen_close" ) 
		{
            //@ts-ignore
            PSPDFKit.unload(".pspdfkit-container");
			this._container;
			this.div.style.height = "10%";
			this.div.style.width = "10%"; 
			this.notes.style.display = "block";
        }
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
    }
    private async RenderThumbnails(files: AttachedFile[], container: HTMLDivElement) {
        var elements:any = document.getElementsByClassName("card");
                            while(elements.length > 0){
                                console.log("removeNode")
                            elements[0].parentNode.removeChild(elements[0]);
                        }
		for (let index = 0; index < files.length; index++) {			

            this.newCard = document.createElement("div");
			const file = files[index];

            this.newCard.classList.add("card");
			this.newCard.innerHTML = `
					<h3>${file.subject}</h3>
                	<h3>${file.fileName}</h3>
                `;
                container.appendChild(this.newCard);
				this.btn = document.createElement("Button");
                this.btn.id = file.annotationId;
                this.btn.innerHTML = "Edit PDF";
                this.btn.classList.add("btn");
                this.btn.onclick = () =>this.PSPDFKit(file);
				this.newCard.appendChild(this.btn);
                container.appendChild(this.newCard);              			
		}
	}
    private async GetFiles(ref: EntityReference): Promise<AttachedFile[]> {
		console.log(ref);
		let fetchXml =
            "<fetch>"+
                "<entity name='annotation'>"+
                    "<filter>"+
                        "<condition attribute='objectid' operator='eq' value='"+ref.id+"' />"+
                    "</filter>"+
                        "<filter type='and'>"+
                            "<condition attribute='mimetype' operator='eq' value='application/pdf' />"+
                    "</filter>"+
                "</entity>"+
            "</fetch>";          
		let query = '?fetchXml=' + encodeURIComponent(fetchXml);

		try {
            const result = await this._context.webAPI.retrieveMultipleRecords("annotation", query);
			let items: AttachedFile[] = [];
			for (let i = 0; i < result.entities.length; i++) {
				let record = result.entities[i];
				let annotationId = <any>record["annotationid"];
				let fileName = <string>record["filename"];
				let subject = <string>record["subject"];
				let description = <string>record["description"];
				let mimeType = <string>record["mimetype"];
				let content = <string>record["body"] || <string>record["documentbody"];
				let fileSize = <number>record["filesize"];
				const ext = fileName.substr(fileName.lastIndexOf('.')).toLowerCase();
				let file = new AttachedFile( annotationId, fileName, mimeType, content, fileSize, subject, description);
				items.push(file);
			}
			return items;
		}
		catch (error) {

			return [];
		}
    }

    public convertBase64ToArrayBuffer(file: any):ArrayBuffer {
        var base64result = file;
        let binaryString  = window.atob(base64result);
		const len = binaryString.length;
		const buffer = new ArrayBuffer(len);

        const bytes = new Uint8Array(buffer);
        for (var i = 0; i < len; i++) {
           bytes[i] = binaryString.charCodeAt(i);
        }       
        return bytes.buffer;
    }

	private convertArrayBufferToBase64(buffer: ArrayBuffer): string {
        let binary = "";
        const bytes = new Uint8Array(buffer);
        const len = bytes.length;
        for (var i = 0; i < len; i++) {
            binary += String.fromCharCode(bytes[i]);
        }
        return btoa(binary);
    }
	public async PSPDFKit(file: AttachedFile)
    {      
        var annoId:any;
		this._context.mode.setFullScreen(true);
        const PSPDFKit = await import("../node_modules/pspdfkit/");
        //@ts-ignore
        PSPDFKit.unload(".pspdfkit-container");
		this.div.style.height = "100vh";
        this.div.style.width = "100%"; 
		this.notes.style.display = "none";
		this.arrBuffer = this.convertBase64ToArrayBuffer(file.fileContent);

        //@ts-ignore
        PSPDFKit.load({
            disableWebAssemblyStreaming: true,
            baseUrl : "https://pspdfkitstorage1.blob.core.windows.net/$web/distdynamics/",
            licenseKey: this._context.parameters.psPdfKitLicenseKey.raw,
            container: ".pspdfkit-container",
			toolbarItems: [
                //@ts-ignore
				...PSPDFKit.defaultToolbarItems,
				{ type: "content-editor" }
			],
			document: this.arrBuffer,                       
        }).then((instance: any) => {
            this._instance = instance;
            const saveButton = {
                type: "custom",
                id: "btnSavePdf",
                title: "Save",
                onPress: async (event: any) => {
                    console.log('add' +this._container.parentElement?.parentElement);
                    
                    this._container.classList.add('loading');
                    this.overlayDiv= document.getElementsByClassName('symbolFont Close-symbol')[1];
                    if(this.overlayDiv){
                        this.overlayDiv.parentElement.style.display = "none"
                    }
                    // export pdf
                    const pdfBuffer = await this._instance.exportPDF();
                    // convert pdf to base64 string
                    const pdfBase64 = this.convertArrayBufferToBase64(pdfBuffer);
                    // update annotation entity
                    this._context.webAPI.updateRecord("annotation", file.annotationId, { "documentbody": pdfBase64, "mimetype": "application/pdf" })                      
                        .then((val) => {                          
                            annoId = val.id;   
                        }, (reason) => {
                            this._context.navigation.openErrorDialog({ message: "Failed to Save PDF!" });
                        })
                        .finally(() => {  
                            this.RefreshThumbnail(annoId,pdfBase64);                                    
                        });					
                }
            }
            // add save button
            instance.setToolbarItems((items: { type: string; id: string; title: string; onPress: (event: any) => void; }[				
			]) => {				
                items.push(saveButton);
                return items;
            });
        })
       
        .catch(console.error);		
    }

    private async RefreshThumbnail(annoID:any, base64:any)
    {
        let reference: EntityReference = new EntityReference(
            (<any>this._context).page.entityTypeName,
            (<any>this._context).page.entityId
        )
        let fetchXml =
            "<fetch>"+
                "<entity name='annotation'>"+
                    "<filter>"+
                        "<condition attribute='annotationid' operator='eq' value='"+annoID+"' />"+
                    "</filter>"+
                        "<filter type='and'>"+
                            "<condition attribute='mimetype' operator='eq' value='application/pdf' />"+
                    "</filter>"+
                "</entity>"+
            "</fetch>";          
		let query = '?fetchXml=' + encodeURIComponent(fetchXml);
        let items: AttachedFile[] = [];
		try {
            const result = await this._context.webAPI.retrieveMultipleRecords("annotation", query);			
			for (let i = 0; i < result.entities.length; i++) {
				let record = result.entities[i];
				let annotationId = <any>record["annotationid"];
				let fileName = <string>record["filename"];
				let subject = <string>record["subject"];
				let description = <string>record["description"];
				let mimeType = <string>record["mimetype"];
				let content = <string>record["body"] || <string>record["documentbody"];
				let fileSize = <number>record["filesize"];
				const ext = fileName.substr(fileName.lastIndexOf('.')).toLowerCase();
				let file = new AttachedFile( annotationId, fileName, mimeType, content, fileSize, subject, description);
				items.push(file);
			}
		}
		catch (error) {
            return [];
		}
        if(items){
            var editButton = document.getElementById(annoID);          
            if(editButton){
                editButton.onclick = () =>this.PSPDFKit(items[0]);     
                this._container.classList.remove('loading');
                this.overlayDiv= document.getElementsByClassName('symbolFont Close-symbol')[1];
                    if(this.overlayDiv){
                        this.overlayDiv.parentElement.style.display = "block"
                }                     
            }
        }
    }

}
