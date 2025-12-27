import { Component } from '@angular/core';
import { NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { HttpClient } from '@angular/common/http';
import { BlockUiDialog } from '../dialogs/block-ui-dialog';
import { ErrorMessageDialog } from '../dialogs/error-message-dialog';


let _documentEditorDemoComponent: DocumentEditorDemoComponent;


@Component({
  selector: 'document-editor-demo',
  templateUrl: './document-editor-demo.html',
})
export class DocumentEditorDemoComponent {

  // Document editor.
  _documentEditor: Vintasoft.Imaging.Office.UI.WebDocumentEditorJS | null = null;

  // Dialog that allows to block UI.
  _blockUiDialog: BlockUiDialog | null = null;



  constructor(public modalService: NgbModal, private httpClient: HttpClient) {
    _documentEditorDemoComponent = this;
  }



  /**
   * Component is initializing.
   */
  ngOnInit() {
    // get identifier of current HTTP session
    this.httpClient.get<any>('api/Session/GetSessionId').subscribe(data => {
      // set the session identifier
      Vintasoft.Shared.WebImagingEnviromentJS.set_SessionId(data.sessionId);

      // specify web services, which should be used by Vintasoft Web Document Editor  ("defaultImageCollectionService" and "defaultImageService" are necessary for printing functionality only)

      Vintasoft.Shared.WebServiceJS.defaultFileService = new Vintasoft.Shared.WebServiceControllerJS("vintasoft/api/MyVintasoftFileApi");
      Vintasoft.Shared.WebServiceJS.defaultOfficeService = new Vintasoft.Shared.WebServiceControllerJS("vintasoft/api/MyVintasoftOfficeApi");
      Vintasoft.Shared.WebServiceJS.defaultImageCollectionService = new Vintasoft.Shared.WebServiceControllerJS("vintasoft/api/MyVintasoftImageCollectionApi");
      Vintasoft.Shared.WebServiceJS.defaultImageService = new Vintasoft.Shared.WebServiceControllerJS("vintasoft/api/MyVintasoftImageApi");

      // create settings for web document editor
      let documentEditorSettings: Vintasoft.Imaging.Office.UI.WebDocumentEditorSettingsJS =
        new Vintasoft.Imaging.Office.UI.WebDocumentEditorSettingsJS("documentEditorContainer", "documentEditor");

      // create web document editor
      this._documentEditor = new Vintasoft.Imaging.Office.UI.WebDocumentEditorJS(documentEditorSettings);

      // subscribe to the "warningOccured" event of document editor
      Vintasoft.Shared.subscribeToEvent(this._documentEditor, "warningOccured", this.__documentEditor_warningOccured);
      // subscribe to the "formulaSyntaxError" event of document editor
      Vintasoft.Shared.subscribeToEvent(this._documentEditor, "synchronizationException", this.__documentEditor_synchronizationException);
      // subscribe to the asyncOperationStarted event of document editor
      Vintasoft.Shared.subscribeToEvent(this._documentEditor, "asyncOperationStarted", this.__documentEditor_asyncOperationStarted);
      // subscribe to the asyncOperationFinished event of document editor
      Vintasoft.Shared.subscribeToEvent(this._documentEditor, "asyncOperationFinished", this.__documentEditor_asyncOperationFinished);
      // subscribe to the asyncOperationFailed event of document editor
      Vintasoft.Shared.subscribeToEvent(this._documentEditor, "asyncOperationFailed", this.__documentEditor_asyncOperationFailed);
      // subscribe to the "saveChangesRequest" event of document editor
      Vintasoft.Shared.subscribeToEvent(this._documentEditor, "saveChangesRequest", this.__documentEditor_saveChangesRequest);

      // open the default document
      this.__openDefaultDocument();
    });
  }




  // === Document editor events ===

  __documentEditor_warningOccured(event: any, eventArgs: any) {
    // show the error message
    _documentEditorDemoComponent.__showErrorMessage(eventArgs.message);
  }

  __documentEditor_synchronizationException(event: any, eventArgs: any) {
    // show the error message
    _documentEditorDemoComponent.__showErrorMessage(eventArgs.message);
  }

  __documentEditor_asyncOperationStarted(event: any, eventArgs: any) {
    // block UI
    _documentEditorDemoComponent.__blockUI(eventArgs.description);
  }

  __documentEditor_asyncOperationFinished(event: any, eventArgs: any) {
    // unblock UI
    _documentEditorDemoComponent.__unblockUI();
  }

  __documentEditor_asyncOperationFailed(event: any, eventArgs: any) {
    // unblock UI
    _documentEditorDemoComponent.__unblockUI();

    // get description of asynchronous operation
    var description = eventArgs.description;
    // get additional information about asynchronous operation
    var additionalInfo = eventArgs.data;
    // if additional information exists
    if (additionalInfo != null)
      // show error message
      _documentEditorDemoComponent.__showErrorMessage(additionalInfo);
    // if additional information does NOT exist
    else
      // show error message
      _documentEditorDemoComponent.__showErrorMessage(description + ": unknown error.");
  }

  __documentEditor_saveChangesRequest(event: any, eventArgs: any) {
    if (!confirm("Document is changed and needs to be saved. Do you want to save document?")) {
      eventArgs.cancel = true;
    }
  }


  // === Open default document ===

  __openDefaultDocument() {
    var fileId = "DocxTestDocument.docx";
    // copy the file from global folder to the session folder
    Vintasoft.Imaging.VintasoftFileAPI.copyFile("UploadedImageFiles/" + fileId, _documentEditorDemoComponent.__onCopyFile_success, _documentEditorDemoComponent.__onCopyFile_error);
  }

  /**
   Request for copying of file is executed successfully.
   @param {object} data Information about copied file.
  */
  __onCopyFile_success(data: any) {
    if (_documentEditorDemoComponent._documentEditor != null) {
      // open document in the document editor
      _documentEditorDemoComponent._documentEditor.openDocument(data.fileId);
    }
  }

  /**
   Request for copying of file is failed.
   @param {object} data Information about error.
  */
  __onCopyFile_error(data: any) {
    alert(data.errorMessage);
  }


  // === Utils ===

  /**
   * Blocks the UI. 
   * @param text Message that describes why UI is blocked.
   */
  __blockUI(text: string) {
    _documentEditorDemoComponent._blockUiDialog = new BlockUiDialog(_documentEditorDemoComponent.modalService);
    _documentEditorDemoComponent._blockUiDialog.message = text;
    _documentEditorDemoComponent._blockUiDialog.open();
  }

  /**
   Unblocks the UI.
  */
  __unblockUI() {
    if (_documentEditorDemoComponent._blockUiDialog != null && _documentEditorDemoComponent._blockUiDialog !== undefined)
      _documentEditorDemoComponent._blockUiDialog.close();
  }

  /**
   * Shows an error message.
   * @param data Information about error.
   */
  __showErrorMessage(data: any) {
    _documentEditorDemoComponent.__unblockUI();
    let dlg: ErrorMessageDialog = new ErrorMessageDialog(_documentEditorDemoComponent.modalService);
    dlg.errorData = data;
    dlg.open();
  }

  /**
   Returns a value indicating whether application is executing on mobile device.
  */
  __isMobileDevice() {
    const toMatch = [
      /Android/i,
      /webOS/i,
      /iPhone/i,
      /iPad/i,
      /iPod/i,
      /BlackBerry/i,
      /Windows Phone/i
    ];

    return toMatch.some((toMatchItem) => {
      return navigator.userAgent.match(toMatchItem);
    });
  }

}
