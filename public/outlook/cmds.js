console.log("cmds");
let thegoods = "thegoods";
let mailboxItem;

Office.initialize = function (reason) {};

Office.onReady(() => {
  document.getElementById(thegoods).innerHTML = "On Ready complete </br>";

  try {
    mailboxItem = Office.context.mailbox.item;

    document.getElementById(thegoods).innerHTML += "</br>" + JSON.stringify(mailboxItem);
  } catch (error) {}
});

function validateOnSend(eventArgs) {
  console.log("validateOnSend", eventArgs);
  //   // Prevent the item from being sent immediately
  //   eventArgs.completed({ allowEvent: false });
  //  // Allow the item to be sent after the asynchronous operation completes
  //  eventArgs.completed({ allowEvent: true });

  // console.log(Office.context);
  console.log(Office.context.mailbox.item);
  if (!Office.context.mailbox.item.itemId) {
    Office.context.mailbox.item.saveAsync(function (data) {
      let id = data.value;
      console.log(id);
      console.log(Office.context.mailbox.item.itemId);
      setTimeout(() => eventArgs.completed({ allowEvent: true }), 1000);
    });
  } else {
    eventArgs.completed({ allowEvent: true });
  }
}
