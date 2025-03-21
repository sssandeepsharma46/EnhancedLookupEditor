import { it } from "node:test";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { fetchLookupRecord } from "./helper/dataversehelper";

interface lookupRec {
    lookupId: string;
    name: string;
    columnName?: string;
}

export class editlookup implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _lookupNameInput: HTMLInputElement;
    private _lookupColumnInput: HTMLInputElement;
    private _editButton: HTMLButtonElement;
    private _saveButton: HTMLButtonElement;
    private _selectLookupButton: HTMLButtonElement;
    private _deleteButton: HTMLButtonElement;
    private _lookupId: string | null;
    private _lookupSearchInput: HTMLInputElement;
    private _lookupList: HTMLUListElement;
    private _lookupListContainer: HTMLDivElement | null = null;
    private _lookupSchemaName: string | "";
    private _lookupColumnSchemaName: string | "";
    private _parentRecordId: string | "";
    private _parentRecordSchemaname: string | "";
    private _parentEntityBindSchema: string | "";
    private _lookupEntityNameColumnSchema: string | "";
    private _notifyOutputChanged: () => void;
    /**
     * Empty constructor.
     */
    constructor() {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        // Add control initialization code
        this._context = context;
        this._container = container;
        this._notifyOutputChanged = notifyOutputChanged;
        this._lookupId = context.parameters.lookupFieldId.raw && context.parameters.lookupFieldId.raw.length > 0 ? context.parameters.lookupFieldId.raw[0].id : null;
        this._lookupSchemaName = context.parameters.lookupEntitySchemaName.raw && context.parameters.lookupEntitySchemaName.raw.length > 0 ? context.parameters.lookupEntitySchemaName.raw : "";
        this._lookupEntityNameColumnSchema = context.parameters.lookupEntityNameColumnSchema.raw && context.parameters.lookupEntityNameColumnSchema.raw.length > 0 ? context.parameters.lookupEntityNameColumnSchema.raw : "";
        this._lookupColumnSchemaName = context.parameters.columnsOfLookupEntity.raw && context.parameters.columnsOfLookupEntity.raw.length > 0 ? context.parameters.columnsOfLookupEntity.raw : "";
        this._parentRecordSchemaname = context.parameters.parentEntitySchemaName.raw && context.parameters.parentEntitySchemaName.raw.length > 0 ? context.parameters.parentEntitySchemaName.raw : "";
        this._parentEntityBindSchema = context.parameters.parentEntityBindSchema.raw && context.parameters.parentEntityBindSchema.raw.length > 0 ? context.parameters.parentEntityBindSchema.raw : "";
        this._parentRecordId = context.parameters.parentRecordId.raw && context.parameters.parentRecordId.raw.length > 0 ? context.parameters.parentRecordId.raw : "";
        //create DOM elements
        this._lookupNameInput = document.createElement("input");
        this._lookupColumnInput = document.createElement("input");
        this._editButton = document.createElement("button");
        this._saveButton = document.createElement("button");
        this._selectLookupButton = document.createElement("button");
        this._deleteButton = document.createElement("button");

        this._editButton.innerText = "Edit";
        this._saveButton.innerText = "Save";
        this._selectLookupButton.innerText = "Select Record";
        this._deleteButton.innerText = "Delete";
        this._saveButton.style.display = "none";

        this._lookupNameInput.classList.add("lookup-details-card", "input");
        this._lookupColumnInput.classList.add("lookup-details-card", "input");
        this._editButton.classList.add("lookup-details-card", "button", "edit-button");
        this._saveButton.classList.add("lookup-details-card", "button", "save-button");
        this._selectLookupButton.classList.add("lookup-details-card", "select-lookup-button", "button");
        this._deleteButton.classList.add("lookup-details-card", "close-button");

        this._editButton.addEventListener("click", this.onEdit.bind(this));
        this._saveButton.addEventListener("click", this.onSave.bind(this));
        this._selectLookupButton.addEventListener("click", this.onSelectLookup.bind(this));
        this._deleteButton.addEventListener("click", this.onDelete.bind(this));

        this._container.appendChild(this._lookupNameInput);
        this._container.appendChild(this._lookupColumnInput);
        this._container.appendChild(this._editButton);
        this._container.appendChild(this._saveButton);
        this._container.appendChild(this._selectLookupButton);
        this._container.appendChild(this._deleteButton);

        this.renderLookupDetails();
    }

    private onSelectLookup(): void {
        if (!this._lookupListContainer) {
            this._lookupListContainer = document.createElement("div");
            this._lookupListContainer.classList.add("lookup-container");

            this._lookupSearchInput = document.createElement("input");
            this._lookupSearchInput.type = "text";
            this._lookupSearchInput.placeholder = "Search your records...";
            this._lookupListContainer.appendChild(this._lookupSearchInput);

            this._lookupList = document.createElement("ul");
            this._lookupList.classList.add("lookup-list");
            this._lookupListContainer.appendChild(this._lookupList);

            this._container.appendChild(this._lookupListContainer);
            this._lookupSearchInput.addEventListener("input", this.onSearchLookup.bind(this, this._lookupSearchInput, this._lookupList));


        }
    }

    private async onSearchLookup(inputField: HTMLInputElement, resultsList: HTMLUListElement): Promise<void> {
        const searchQuery = inputField.value.trim();

        if (searchQuery.length > 2) {
            const lookups = await this.fetchLookups(searchQuery);
            resultsList.innerHTML = "";
            lookups.forEach(item => {
                const listItem = document.createElement("li");
                listItem.textContent = item.name;
                listItem.classList.add("lookup-item");

                listItem.addEventListener("click", () => { this.onLookupSelect(item) });
                resultsList.appendChild(listItem);
            });
        }
    }

    private async fetchLookups(searchQuery: string): Promise<lookupRec[]> {
        const query = `?$select=${this._lookupEntityNameColumnSchema},${this._lookupColumnSchemaName}&$filter=startswith(${this._lookupEntityNameColumnSchema},'${searchQuery}')`;
        const lookupsRecords = await this._context.webAPI.retrieveMultipleRecords(this._lookupSchemaName, query);
        return lookupsRecords.entities.map(entity => ({
            lookupId: entity[`${this._lookupSchemaName}id`],
            name: entity[this._lookupEntityNameColumnSchema],
            lookupColumn: entity[this._lookupColumnSchemaName],
        }));
    }

    private onLookupSelect(lookup: lookupRec): void {
        this._lookupId = lookup.lookupId;

        if (this._lookupListContainer) {
            this._lookupListContainer.style.display = "none";
        }
        // Update the parentcustomerid field on the Contact record
        this.updateLookupRecord(lookup.lookupId);

        // Render the account details after selection
        this.renderLookupDetails();
        this._notifyOutputChanged();
    }

    private async updateLookupRecord(accountId: string): Promise<void> {
        try {
            // Get the current Contact record ID from the context (the contact ID that this control is bound to)
            const parentRecId = this._context.parameters.parentRecordId.raw; // This will give the Contact ID as a string

            if (!parentRecId) {
                console.error("No Contact ID found");
                return;
            }
            this._parentRecordId = parentRecId;
            // Update the parentcustomerid lookup field
            // Construct the update data for parentcustomerid lookup field
            const updateData = {
                //"parentcustomerid_account@odata.bind": `/accounts(${accountId})`, // Use the @odata.bind annotation to specify the URL of the related record
                [`${this._parentEntityBindSchema}@odata.bind`]: `/${this._lookupSchemaName}s(${this._lookupId})`
            };

            // Perform the update using the web API
            await this._context.webAPI.updateRecord(this._parentRecordSchemaname, parentRecId, updateData);
        } catch (error) {
            console.error("Error updating parent lookup value:", error);
        }
    }

    private async renderLookupDetails(): Promise<void> {
        if (this._lookupId) {
            const lookupRecord = await fetchLookupRecord(this._context.webAPI, this._lookupId, this._lookupSchemaName, this._lookupColumnSchemaName, this._lookupEntityNameColumnSchema);
            if (lookupRecord) {
                this._lookupNameInput.value = lookupRecord.name;
                this._lookupColumnInput.value = lookupRecord.columnName;
            }
            this._lookupNameInput.disabled = true;
            this._lookupColumnInput.disabled = true;

            this._editButton.style.display = "inline";
            this._saveButton.style.display = "none";
            this._selectLookupButton.style.display = "none";
        } else {
            this._lookupNameInput.value = "";
            this._lookupColumnInput.value = "";
            this._lookupNameInput.disabled = false;
            this._lookupColumnInput.disabled = false;

            this._editButton.style.display = "none";
            this._saveButton.style.display = "none";
            this._selectLookupButton.style.display = "inline";

            this._deleteButton.style.display = this._lookupNameInput ? "inline" : "none";
            this._selectLookupButton.addEventListener("click", this.onSelectLookup.bind(this));
        }
    }

    private onEdit(): void {
        this._lookupNameInput.disabled = false;
        this._lookupColumnInput.disabled = false;
        this._editButton.style.display = "none";
        this._saveButton.style.display = "inline";
    }

    private async onSave(): Promise<void> {
        if (this._lookupId) {
            try {
                const columnLookupName = this._lookupColumnSchemaName;
                const nameLookupColumn = this._lookupEntityNameColumnSchema;

                if (!this._lookupNameInput.value || !this._lookupColumnInput.value) {
                    console.error("Both fields must be filled before saving.");
                    return;
                }

                // Ensure the update record is type-safe
                const updateRecord: Record<string, unknown> = {
                    [nameLookupColumn]: this._lookupNameInput.value,
                    [columnLookupName]: this._lookupColumnInput.value,
                };
                await this._context.webAPI.updateRecord(this._lookupSchemaName, this._lookupId, updateRecord);
                this._editButton.style.display = "inline";
                this._saveButton.style.display = "none";
                this._lookupNameInput.disabled = true;
                this._lookupColumnInput.disabled = true;
            } catch (error) {
                console.error("Error on save : ", error);
                return;
            }
        }
    }

    private async onDelete(): Promise<void> {
        this._lookupId = null;
        try {
            const parentRecordId = this._context.parameters.parentRecordId.raw;
            if (!parentRecordId) {
                console.error("No parent record id found.");
                return;
            }

            const updateData = {
                // Use the @odata.bind annotation to specify the URL of the related record
                [`${this._parentEntityBindSchema}@odata.bind`]: null
            };

            // Perform the update using the web API
            await this._context.webAPI.updateRecord(this._parentRecordSchemaname, parentRecordId, updateData);

            this.renderLookupDetails();
            this._notifyOutputChanged();

        } catch (error) {
            console.error("Error updating lookup : ", error);
        }
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Add code to update control view
        this._context = context;
        this._lookupId = context.parameters.lookupFieldId.raw && context.parameters.lookupFieldId.raw.length > 0 ? context.parameters.lookupFieldId.raw[0].id : null;
        this.renderLookupDetails();
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
        this._editButton.removeEventListener("click", this.onEdit.bind(this));
        this._saveButton.removeEventListener("click", this.onSave.bind(this));
        this._selectLookupButton.removeEventListener("click", this.onSelectLookup.bind(this));
        this._deleteButton.removeEventListener("click", this.onDelete.bind(this));
    }
}
