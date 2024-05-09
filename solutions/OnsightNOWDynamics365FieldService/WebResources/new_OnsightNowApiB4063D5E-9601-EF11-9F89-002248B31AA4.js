'use strict';

// depending on where this script is utilized (e.g. within an HTML web resource),
// we may have to obtain the Xrm object from the parent
var xrm = typeof(Xrm) !== "undefined" ? Xrm : parent.Xrm;
var globalContext = xrm ? xrm.Utility.getGlobalContext() : {};
var clientContext = globalContext ? globalContext.client : {};

// Regex for Xrm.WebApi calls returning an array of entities
const WebApiExpandExpr = /(\w+)\(\$select\=(\w+)\)$/;

/**
 * The following 'x2y' structures map a D365 entity to a child field within that
 * entity which either:
 *     a) is the email address of the user associated with that entity, or
 *     b) leads to the next entity in the "chain", which ultimately leads to an email address
 */
const SystemUser2Email = [
    {
        entityLogicalName: "systemuser",
        select: "internalemailaddress"
    }
];

const BookableResource2Email = [
    {
        entityLogicalName: "bookableresource",
        select: "_userid_value"
    },
    ...SystemUser2Email
];

const BookableResourceBooking2Email = [
    {
        entityLogicalName: "bookableresourcebooking",
        select: "_resource_value"
    },
    ...BookableResource2Email
];

const WorkOrder2FieldTechEmail = [
    {
        entityLogicalName: "msdyn_workorder",
        expand: "msdyn_msdyn_workorder_bookableresourcebooking_WorkOrder($select=bookableresourcebookingid)"
    },
    ...BookableResourceBooking2Email
];

const WorkOrder2RemoteExpertEmail = [
    {
        entityLogicalName: "msdyn_workorder",
        select: "_msdyn_supportcontact_value"
    },
    ...BookableResource2Email
];

/**
 * Uses a mapping struct (such as an element of EmailMappings) to extract a value
 * from the given recordResult.
 * @param {*} recordResult 
 * @param {*} recordMapping 
 */
function mapResultToValue(recordResult, recordMapping) {
    // If the mapping defines an "expand" property, it means we expect an array result
    if (recordMapping.expand) {
        // Result must be an array
        let matches = recordMapping.expand.match(WebApiExpandExpr);
        if (matches && matches.length === 3) {
            let arrName = matches[1];
            let elemName = matches[2];

            // Grab the first element in the array (if available) and map the result value
            // from the expand expression's inner "$select".
            let arr = recordResult[arrName];
            if (Array.isArray(arr) && arr.length > 0) {
                return arr[0][elemName];
            }
        } 
    }
    else if (recordMapping.select) {
        // Result must be a single value
        return recordResult[recordMapping.select];
    }

    // If we get here, we have no explicit way of mapping to a result value, so just return the resultRecord
    return recordResult;
}

/**
 * Retrieve an entity value based on the given initial entity name and ID,
 * along with a mappings object which directs how the record retrieval is
 * mapped to a result value.
 * 
 * Use this method to chain together record retrievals when the result value
 * is referenced by another entity (the starting point).
 * @param {[*]} mappings 
 * @param {string} entityId 
 * @returns {Promise<any>}
 */
async function retrieveRecordAsync(mappings, entityId) {
    let mapping = mappings[0];
    let options = "";

    if (mapping.select) {
        options += (options.length === 0) ? "?" : "&";
        options += "$select=" + mapping.select;
    }
    if (mapping.expand) {
        options += (options.length === 0) ? "?" : "&";
        options += "$expand=" + mapping.expand;
    }

    const entityLogicalName = mapping.entityLogicalName;

    let result = await Xrm.WebApi.retrieveRecord(entityLogicalName, entityId, options);
    let resultValue = mapResultToValue(result, mapping);

    if (mappings.length === 1) {
        return resultValue;
    }

    return await retrieveRecordAsync(mappings.slice(1), resultValue);
}

/**
 * Navigate to the given URL using the Dynamics Xrm.Navigation API
 * @param {string} url URL to navigate to
 */
function navigateToUrl(url) {
    if (xrm) {
        xrm.Navigation.openUrl(url);
    }
}

/**
 * Get the D365 entity which triggered the button click event. This could be the primary
 * entity (if the main form command button was clicked) or a sub-entity if the button
 * is located within a subgrid (such as the Bookable Resources subgrid section within
 * the Work Order main form).
 * @param {*} primaryControl 
 */
function getTriggeringEntity(primaryControl) {
    // If there's a grid associated w/the given control, we have a list
    // of entities and must pick the first selected one.
    const grid = primaryControl.getGrid && primaryControl.getGrid();
    if (grid) {
        const selectedEntities = grid.getSelectedRows().getAll();
        return (selectedEntities.length === 0) ? null : selectedEntities[0];
    }

    // Otherwise we're dealing with a top-level (MainForm) entity
    return primaryControl.data.entity;
}

/**
 * Get the email address associated with the given entity. For any entity other than systemuser,
 * we will drill down into "child" entities until we find the corresponding systemuser and their email.
 * 
 * The entityType may be one of:
 *      msdyn_workorder, bookableresourcebooking, bookableresource, or systemuser
 * @param {string} entityType 
 * @param {string} entityId 
 * @param {string} callTargetType "fieldtech" | "expert" | null | undefined
 * @returns {Promise<string>} The email address.
 */
function getEmailAddressAsync(entityType, entityId, callTargetType) {
    let mappings = [];
    switch (entityType) {
        case "msdyn_workorder":
            mappings =
                callTargetType === "fieldtech"
                    ? WorkOrder2FieldTechEmail
                    : WorkOrder2RemoteExpertEmail;
            break;
        case "bookableresourcebooking":
            mappings = BookableResourceBooking2Email;
            break;
        case "bookableresource":
            mappings = BookableResource2Email;
            break;
        case "systemuser":
            mappings = SystemUser2Email;
            break;
    }

    return retrieveRecordAsync(mappings, entityId);
}

function execScheduleMeetingAction(participantEmails) {
    var parameters = {};
    parameters.ParticipantEmails = participantEmails;

    var req = new XMLHttpRequest();
    req.open("POST", Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/new_ScheduleNOWMeeting", true);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.setRequestHeader("Accept", "application/json");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status >= 200 && this.status <= 204) {
                var result = JSON.parse(this.response);
                console.log(result);
                // Return Type: mscrm.new_ScheduleNOWMeetingResponse
                // Output Parameters
                var meetingUrl = result["MeetingUrl"];
                navigateToUrl(meetingUrl);
            } else {
                console.log(this.responseText);
            }
        }
    };
    req.send(JSON.stringify(parameters));
}

/**
 * Launches Onsight NOW in a new tab by first creating a meeting with the designated target.
 * @param {any} primaryControl the D365 UI element which triggered this handler.
 * @param {bool} callTargetType the command parameter passed in from the UI. For Work Orders,
 * this will be either "fieldtech" or "expert".
 * @return none
 */
async function onClick(primaryControl, callTargetType) {
    console.log("onClick");

    // Get the entity in context (could be a selected entity within a subgrid, for example)
    const contextEntity = getTriggeringEntity(primaryControl);
    if (!contextEntity) {
        // Skipping launch; no entity in context
        return;
    }
    
    // Assume that the current D365 user is the person initiating the Onsight NOW meeting
    const callerEmail = await getEmailAddressAsync("systemuser", xrm.Page.context.getUserId());
    const calleeEmail = await getEmailAddressAsync(contextEntity._entityType, contextEntity._entityId.guid, callTargetType)

    const participantEmails = `${callerEmail},${calleeEmail}`;
    console.log(`Scheduling NOW meeting with participants ${participantEmails}`);
    execScheduleMeetingAction(participantEmails);
}
