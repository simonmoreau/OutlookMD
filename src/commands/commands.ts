/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
async function action(event: Office.AddinCommands.Event) {
  const item = Office.context.mailbox.item as Office.MessageRead | Office.AppointmentRead;

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

  // Type assertion for AppointmentRead
  const appointment = item as Office.AppointmentRead;

  // Retrieve appointment details
  const subject: string = await getValue<string>(item.subject);
  const start = await getValue<Date>(item.start);
  const end = await getValue<Date>(item.end);
  const location = await getValue<string>(item.location);
  const requiredAttendees: Office.EmailAddressDetails[] = await getValue<Office.EmailAddressDetails[]>(appointment.requiredAttendees) ?? [];
  const optionalAttendees: Office.EmailAddressDetails[] = await getValue<Office.EmailAddressDetails[]>(appointment.optionalAttendees) ?? [];
  const organizer: Office.EmailAddressDetails | undefined = await getValue<Office.EmailAddressDetails | undefined>(appointment.organizer);

  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: subject + " has been selected." + end.toDateString(),
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);

/**
 * Gets the value of a property from an Office object asynchronously.
 * @param input The Office object to get the value from.
 * @returns A promise that resolves with the value of the property.
 */
function getValue<T>(input : any): Promise<T> {
      return new Promise<T>((resolve, reject) => {
            input.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            reject(asyncResult.error.message);
            return;
        }

        // Display the subject on the page.
        resolve(asyncResult.value);
    });
  });
}
