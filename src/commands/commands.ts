/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */
import { Liquid } from "liquidjs";

const template = `
# Meeting: {{ subject }}

**Start:** {{ start }}
**End:** {{ end }}
**Location:** {{ location }}

**Organizer:** {{ organizer.displayName }} ({{ organizer.emailAddress }})

**Required Attendees:**
{% for attendee in requiredAttendees %}
- {{ attendee.displayName }} ({{ attendee.emailAddress }})
{% endfor %}

**Optional Attendees:**
{% for attendee in optionalAttendees %}
- {{ attendee.displayName }} ({{ attendee.emailAddress }})
{% endfor %}
`;

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
async function action(event: Office.AddinCommands.Event) {
  const item = Office.context.mailbox.item as Office.AppointmentCompose | Office.AppointmentRead;

  if (item.itemType !== Office.MailboxEnums.ItemType.Appointment) {
    const errorMessage: Office.NotificationMessageDetails = {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: "This command only works on appointments.",
      icon: "Icon.80x80",
      persistent: true,
    };
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "ActionPerformanceNotification",
      errorMessage
    );
    event.completed();
    return;
  }

  let lastError: any = null;

  try {

    // Retrieve appointment details
    const data = await GetAppointementDetails(item);

    const engine = new Liquid();
    const rendered = await engine.parseAndRender(template, data);

    const popup = window.open("", "popup", "width=600,height=400");
    if (popup) {
      popup.document.write(`<pre style="white-space: pre-wrap; margin:0;">${rendered}</pre>`);
      popup.document.close();
    }

    Notify();

  } catch (error) {
    lastError = error;
    // Show error notification
    const errorMessage: Office.NotificationMessageDetails = {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: `Error: ${error.message || error}`,
      icon: "Icon.80x80",
      persistent: true,
    };
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "ActionPerformanceNotification",
      errorMessage
    );
  } finally {

    // Be sure to indicate when the add-in command function is complete.
    event.completed();
  }
}

// Register the function with Office.
Office.actions.associate("action", action);

function Notify() {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "The message has been copied",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );
}

async function GetAppointementDetails(appointment: Office.AppointmentCompose | Office.AppointmentRead) {
  const subject: string = await getValue<string>(appointment.subject);
  const start = await getValue<Date>(appointment.start);
  const end = await getValue<Date>(appointment.end);
  const location = await getValue<string>(appointment.location);
  const requiredAttendees: Office.EmailAddressDetails[] =
    (await getValue<Office.EmailAddressDetails[]>(appointment.requiredAttendees)) ?? [];
  const optionalAttendees: Office.EmailAddressDetails[] =
    (await getValue<Office.EmailAddressDetails[]>(appointment.optionalAttendees)) ?? [];
  const organizer: Office.EmailAddressDetails | undefined = await getValue<
    Office.EmailAddressDetails | undefined
  >(appointment.organizer);

  const data = {
    subject,
    start: start.toLocaleString(),
    end: end.toLocaleString(),
    location,
    organizer,
    requiredAttendees,
    optionalAttendees,
  };
  return data;
}

/**
 * Gets the value of a property from an Office object asynchronously.
 * @param input The Office object to get the value from.
 * @returns A promise that resolves with the value of the property.
 */
function getValue<T>(input: any): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    if (input && typeof input.getAsync === "function") {
          input.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        reject(asyncResult.error.message);
        return;
      }
      // Display the subject on the page.
      resolve(asyncResult.value);
    });
    } else {
      resolve(input as T);
    }
  });
}
