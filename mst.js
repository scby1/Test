/*
For this script to work you must be viewing Teams from your browser.
https://teams.microsoft.com
This script was tested in Firefox.
You must be using Developer tools of your browser
YOU MUST INSPECT a message, otherwise the HTML elements will not be available
The elements are burried in an iFrame, and you must inspect something for the iFrame code to be available to the script
Once you have inspected something, go to the Console tab of your developer tools
paste this whole script and run it
You can run it by pressing the Play button,
or with the keyboard shortcut Ctrl + Enter
It will first scroll all the way to the first message
Then it will create a new browser window with the messages in it
Then it is up to you to do what you want with it
I suggest you print the content to pdf (Ctrl + P)

It is possible the process will stop before reaching the top
If that is the case and you want to keep scrolling,
you can retrigger the script but call it with a false parameter to maintain the existing storage list
main(false).
*/

/* 
 * how long to wait before the sniffing the message pane after a scroll
 * if it is set too low, or not at all, the messages do not have enough time to load
 * the slower your computer or connection, the higher this value should be
 * the higher this value will be, the longer it will take to process a chat window
 */
var delayTime = 500;
/* log messages to console */
var logToConsole = true;
/* perhaps you only need a few scrolls worth or data */
var maxScrollOccurences = 0;//0 for inifinite
/* 
 * sometimes having the right delay value is not enough
 * so, we add a fudge factor in case the delay was not long enough
 * this allows us to re-sniff the message pane without scrolling again
 * when we run out of fudge, we assume that we have reached the top of the message pane
 */
var maxFudge = () => { return 10; }

/*
you should not need to change any of the variables below
 */


/*
 * This is the div in which the scroll bar resides
 */
var scrollDiv = () => { return document.getElementById("main-window-body").querySelector('[data-tid="message-pane-list-viewport"]'); }
/* 
 * this is where all the messages are displayed
 */
var chatPaneList = () => { return document.getElementById("chat-pane-list"); }
/*
 * all the messages being displayed in the chat pane
 * in a virtual list
 * as more are added, some are dropped and some loose their content
 * so, you must keep track of the messages as you scroll up the message pane
 */
var messagesInChatPane = () => { return chatPaneList().querySelectorAll('.fui-unstable-ChatItem'); }
/* this variable is to stop the scrolling while */
var stop;
/* just a means to keep track of wheter we should keep scrolling or not */
var msgOnSCreenBeforeScroll;
/* just a means to keep track of wheter we should keep scrolling or not */
var msgOnSCreenAfterScroll;
/* this keeps track of how many times we have scrolled so far */
var scrollOccurences = 0;
/* this array will hold the messages to display */
var messagesToDisplay = [];

var firstRender = true;

main();//run this one to start a fresh process
//main(false);//run this one if you want to keep processing an existing window

/*
 * This is the only function call you should be calling
 * set FirstRender to true, or omit it, when you run the script for the first time in a chat pane
 * but, set FirstRender to false for all subsequent runs of the same chat pane
 */
async function main(FirstRender) {
    if (FirstRender != undefined) {
        firstRender = FirstRender
    }
    await ScrollToTopOfMessages();
    PrintIt();
}




async function ScrollToTopOfMessages() {
    var fudgeFactor = maxFudge();
    stop = false;
    msgOnSCreenBeforeScroll = 0;

    if (firstRender) {
        /* start with an empty list */
        messagesToDisplay = [];
    }

    while (!stop) {
        msgOnSCreenBeforeScroll = messagesToDisplay.length;
        LogThis("displayed before: " + msgOnSCreenBeforeScroll);
        await ScrollAndWait();
        scrollOccurences++;
        msgOnSCreenAfterScroll = messagesToDisplay.length;
        LogThis("displayed after: " + msgOnSCreenAfterScroll);

        //if more items are displayed than before, we continue scrolling
        //otherwise we stop and proceed to the next step
        var fudging = false;
        if (msgOnSCreenBeforeScroll == msgOnSCreenAfterScroll) {
            scrollOccurences--;//fudges do no count count. only scrolls that actually occurred
            fudgeFactor--;
            LogThis("fudging " + fudgeFactor + "/" + maxFudge());
            fudging = true;
            if (fudgeFactor == 0) {
                stop = true;
                LogThis("too much fudging");
            }
        }
        else {
            fudgeFactor = maxFudge();
        }

        /*
         * stop scrolling if we have reached the maxScrollOccurences
         */
        if (!fudging && maxScrollOccurences > 0 && scrollOccurences == maxScrollOccurences) {
            LogThis("bumped into maxScrollOccurences");
            stop = true;
        }
    }
}

async function ScrollAndWait() {
    //msgOnSCreenBeforeScroll = msgOnSCreenAfterScroll;
    scrollDiv().scrollTop = 0;
    await delay(delayTime);//let the messages load

    var messages = messagesInChatPane();
    for (var i = messages.length - 1; i > -1; i--) {
        if (messages[i].querySelector('.fui-Divider') != undefined) {
            //we do not want dividers
            LogThis("skipping divider");
            continue;
        }
        var msgId = messages[i].querySelector('[id ^= "timestamp-"]');
        if (msgId == undefined) {
            LogThis("message with no time stamp... ");
            LogThis(messages[i]);
            continue;
        }
        msgId = msgId.id.replace("timestamp-", "");


        if (messagesToDisplay.findIndex(p => p.id === msgId) == -1) {
            LogThis(messagesToDisplay.length + " before adding to messagesToDisplay");
            var en = new entry(messagesToDisplay.length, msgId, messages[i]);
            messagesToDisplay.push(en);
            LogThis(messagesToDisplay.length + " after adding to messagesToDisplay");
        }
    }
}

async function delay(time) {
    return new Promise(resolve => setTimeout(resolve, time));
}

async function PrintIt() {
    LogThis(messagesToDisplay);
    var WinPrint = window.open('', '', 'left=0,top=0,width=800,height=900,toolbar=0,scrollbars=0,status=0');
    for (var i = messagesToDisplay.length - 1; i > -1; i--) {
        var msgDiv = messagesToDisplay[i].element.cloneNode(true);
        msgDiv = RemoveHeaderDiv(msgDiv);
        msgDiv = CleanupEmoji(msgDiv);
        msgDiv = FixImages(msgDiv);
        WinPrint.document.body.appendChild(msgDiv);
    }
    //WinPrint.document.close();
    WinPrint.focus();
    //WinPrint.print();
    //WinPrint.close();
}

/*
 * the header contains the message preview
 * useless in what we are trying to do here
 */
function RemoveHeaderDiv(element) {
    var msgHeader = element.querySelector('[role="heading"]');
    if (msgHeader != undefined) {
        msgHeader.remove();
    }
    return element;
}

/* 
 * the css is not maintained in the new browser window, on purpose
 * and it breaks the Emojis
 * this fixes them
 */
function CleanupEmoji(element) {
    element.querySelectorAll('[itemtype="http://schema.skype.com/Emoji"]').forEach(p => p.setAttribute("src", ""))
    return element;
}

/*
images appear in the new browser window
but most of the images disappear when the window gets printed to pdf
this fixes the problem
 */
function FixImages(element) {
    var imgs = element.querySelectorAll("img");
    for (var i = 0; i < imgs.length; i++) {
        var eleParent = imgs[i].parentElement;
        var tmp = imgs[i].getAttribute("data-orig-src");
        if (tmp != undefined) {
            imgs[i].remove();
            var newimg = document.createElement("img");
            newimg.src = tmp;
            eleParent.appendChild(newimg);
        }
    }
    return element;
}

function LogThis(messageToLog) {
    if (logToConsole) {
        console.log(messageToLog);
    }
}

function entry(index, id, element) {
    this.index = index;
    this.id = id;
    this.element = element;
}










(async () => {
  const expandAll = () => {
    const buttons = document.querySelectorAll('button');
    buttons.forEach(button => {
      if (button.textContent.trim().toLowerCase() === 'see more') {
        button.click();
      }
    });
  };

  const getAllMessages = () => {
    const messageElements = document.querySelectorAll('.ui-chat__message__body, .ui-chat__event__formatted-text');
    let messages = [];
    messageElements.forEach(element => {
       messages.push(element.innerText);
    });
    return messages;
  };
  
  const download = (filename, text) => {
    const element = document.createElement('a');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
    element.setAttribute('download', filename);
  
    element.style.display = 'none';
    document.body.appendChild(element);
  
    element.click();
  
    document.body.removeChild(element);
  };

  expandAll();
  await new Promise(r => setTimeout(r, 2000)); // Wait for "see more" to expand

  const chatMessages = getAllMessages();
  const filename = 'teams_chat_export.txt';
  const text = chatMessages.join('\n---\n');

  download(filename, text);
})();
