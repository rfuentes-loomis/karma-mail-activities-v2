console.log("cmds");
let thegoods = "thegoods";
let mailboxItem;

Office.initialize = function (reason) {
  document.getElementById(thegoods).innerHTML += "</br>initialize complete </br>";

  try {
    mailboxItem = Office.context.mailbox.item;

    document.getElementById(thegoods).innerHTML += "</br>from initialize:</br>" + JSON.stringify(mailboxItem?.itemId);
    document.getElementById(thegoods).innerHTML += "</br>from initialize:</br>" + JSON.stringify(mailboxItem);
    document.getElementById(thegoods).innerHTML += "</br>Mailbox from initialize:</br>" + JSON.stringify(Office.context.mailbox);
  } catch (error) {}
};

Office.onReady(() => {
  document.getElementById(thegoods).innerHTML += "</br>On Ready complete </br>";

  try {
    mailboxItem = Office.context.mailbox.item;

    document.getElementById(thegoods).innerHTML += "</br>Item from onReady:</br>" + JSON.stringify(mailboxItem?.itemId);
    document.getElementById(thegoods).innerHTML += "</br>Item from onReady:</br>" + JSON.stringify(mailboxItem);
    document.getElementById(thegoods).innerHTML += "</br>Mailbox from onReady:</br>" + JSON.stringify(Office.context.mailbox);
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
