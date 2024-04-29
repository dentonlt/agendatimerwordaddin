/** WordWebAddIn_Etude5
 * track traverse agenda item list
 **/

/*eslint max-len: ["error", { "code": 80 }]*/

/** TEST DATA:
0:05 Welcome
0:08 second- item
1 fourth: blah
7 234
15:35 third item (but it's fifth)
6
0:03 final item
 **/

/** global document, Office, Word */

/** Timer Management */
const maxDur: number = 59 * 60 + 59; // max timer (in sec) is 59:59
let curTimeStamp: number = Date.now() / 1000; // optimizer for clock tick
let startTime: number | undefined; // timestamp when timer starts
let timerDuration: number = 0; // duration of timer (in seconds)
let pauseTimeStamp: number | undefined; // used to manage pause action
let runningFlag: number | 0; // 0 = stop, 1 = run, 2 = pause
let currentAgendaItemIndex: number = 0; // which item is running (0-based!)

/** Agenda Item and Duration Management */
/* eslint-disable */

// structure for an agenda item (duration + title)
interface agendaItem {
  duration: number; //duration in seconds
  title: string; // title for item
}

let totalDuration: number = 0; // duration sum of all agenda items
let agendaItems: agendaItem[] = new Array(); // array of agenda items

// regex patterns for capturing agenda items out of the Word Document Body
let regexAgendaItem: RegExp =
  /\b\d+:?\d*\s.+/g;
let regexAgendaTime: RegExp = /\b(\d+:?\d*)/;
let regexAgendaTitle: RegExp = /\d*:?\d{1,2}\s+(.+)/;

/* eslint-enable */

/** HTML/UI elements */
/* eslint-disable */

const body: HTMLBodyElement | null =
  document.getElementById("mainBody") as HTMLBodyElement; // main content

const clockOutput: HTMLDivElement | null =
  document.getElementById("clockOutput") as HTMLDivElement; // wall time
const timerOutput: HTMLDivElement | null =
  document.getElementById("timerOutput") as HTMLDivElement; // timer output

  const startButton: HTMLButtonElement | null =
  document.getElementById("startButton") as HTMLButtonElement; // start button
const startButtonLabel: HTMLSpanElement | null =
  document.getElementById("startButtonLabel") as HTMLSpanElement; // start button
const pauseButton: HTMLButtonElement | null =
  document.getElementById("pauseButton") as HTMLButtonElement; // pause button
const nextButton: HTMLButtonElement | null =
  document.getElementById("nextButton") as HTMLButtonElement; // next item button
const previousButton: HTMLButtonElement | null =
  document.getElementById("previousButton") as HTMLButtonElement; // previous item button

const addOneMinuteButton: HTMLButtonElement | null =
  document.getElementById("addOneMinuteButton") as HTMLButtonElement; // +1min
const subtractOneMinuteButton: HTMLButtonElement | null =
  document.getElementById("subtractOneMinuteButton") as HTMLButtonElement; // -1min
const addFiveMinutesButton: HTMLButtonElement | null =
  document.getElementById("addFiveMinutesButton") as HTMLButtonElement; // +5min
const subtractFiveMinutesButton: HTMLButtonElement | null =
  document.getElementById("subtractFiveMinutesButton") as HTMLButtonElement; // -5min


const reloadButton: HTMLButtonElement | null =
  document.getElementById("reloadButton") as HTMLButtonElement; // read file button
const itemCountLabel: HTMLHeadingElement | null =
  document.getElementById("itemCountLabel") as HTMLHeadingElement; // label for item count
const itemTitleLabel: HTMLHeadingElement | null =
  document.getElementById("itemTitleLabel") as HTMLHeadingElement; // label for item title

/** Styles */
const colorTimeWarning: string = "#CC0000";
const colorTimeNormal: string = "#000000";
/* eslint-enable */

// UTILITIES /////////////////

/** formatTime(): utility function; return input number as a two-digit string */
/**
 * @param time: number. Not clipped, so 0 to big ...
 * @returns string: number as two digits (add leading zero)
 */
function formatTime(time: number): string {
  // convert it to two digits ...
  if (time < 10) return "0" + time;
  else return time.toString();
}

/** durationStringToSeconds(): utility; convert string to seconds */
/**
 * @param durstring: string; convert durstring into seconds
 * @returns number; returns 0 on error (durstring is null or cannot be parsed)
 */
function durationStringToSeconds(durstring: string | null): number {
  if (durstring === null) return 0;

  const substrings: string[] | null = durstring.split(":");

  if (substrings != null) {
    switch (substrings.length) {
      case 1:
        return +substrings[0] * 60;
      case 2:
        // eslint-disable-next-line
        return (+substrings[0] * 60) + (+substrings[1]);
      default:
    }
  }

  return 0;
}

/** remainingTime(): fetch remainingTime value. Avoid typos! */
function remainingTime(timeStamp: number): number {
  if (startTime === undefined) {
    console.log("startTime undefined in remainingTime();");
    return 0;
  }
  return Math.ceil((timerDuration * 1000 - timeStamp + startTime) / 1000);
}

/** setClockText(): update Wall Clock output */
function updateClock(timeobj: Date) {
  if (timeobj === undefined || timeobj === null) {
    console.log("timeobj undefined/null in updateClock()");
    return;
  }
  /* eslint-disable */
  var clockTimeString: string = 
  formatTime(timeobj.getHours()) + ":" +
  formatTime(timeobj.getMinutes()) + ":" +
  formatTime(timeobj.getSeconds());
  /* eslint-enable */
  clockOutput!.textContent = clockTimeString;
}

/** updateTimer(): update timer text as min:sec */
/**
 * @param durationSeconds: number; show this as min:sec on timer */
function updateTimer(durationSeconds: number) {
  if (timerOutput === null) {
    console.log("timerOutput Unitialized in updateTimer()");
    return;
  }

  const min: number = Math.floor((durationSeconds % (60 * 60)) / 60);
  const sec: number = Math.floor(durationSeconds % 60);
  timerOutput.textContent = formatTime(min) + ":" + formatTime(sec);

  // eslint-disable-next-line
  durationSeconds > 0 && durationSeconds < 60
    ? (timerOutput.style.color = colorTimeWarning)
    : (timerOutput.style.color = colorTimeNormal);

  return;
}

/** nextItem(): advance to next agenda item */
function nextItem() {
  currentAgendaItemIndex++;

  // catch out-of-range index, clip
  if (currentAgendaItemIndex >= agendaItems.length) {
    runningFlag = 0;
    currentAgendaItemIndex = agendaItems.length;
    if (itemTitleLabel != null) itemTitleLabel.textContent = "End of Agenda";
    if (itemCountLabel != null) itemCountLabel.textContent = "-";
    updateTimer(0); // using this ensures correct style ...

    // hide next, pause, continue, start buttons
    if (startButton != null) startButton.style.display = "none";
    if (pauseButton != null) pauseButton.style.display = "none";
    if (nextButton != null) nextButton.style.display = "none";
    return;
  }

  // prepare timer update
  timerDuration = agendaItems[currentAgendaItemIndex].duration;
  startTime = Date.now();

  return;
}

/** changeDuration(): add/remove change seconds from running item */
function changeDuration(change: number = 0) {
  if (change === 0) return; // nothing to do

  timerDuration = timerDuration + change;
  if (timerDuration < 1) timerDuration = 1; // minDur = 1 sec
  if (timerDuration > maxDur) timerDuration = maxDur;
}

/* parseAgendaItems(): convert regex finds into dur/title pairs */
/**
 * @param text: string; pull agenda items from here into agendaItems array.
 * @returns number: agendaItems.length
 */
function parseAgendaItems(text: string | null): number {
  agendaItems.length = 0; // clear all items at outset.

  //eslint-disable-next-line
  if ((text === null) || (text.length == 0)) {
    console.log("parseAgendaItems rec'd null/empty text");
    return 0;
  }

  // first, do the matching
  var matches: RegExpMatchArray | null = text.match(regexAgendaItem);
  if (matches === null) {
    console.log("parseAgendaItems() found no matches.");
    return 0;
  }
  /* // DEBUG
  console.log("parseAgendaItems(): found " + matches.length + " items");
  // */

  // second, split all those matches into duration/title pairs
  matches.forEach(function (value: string) {
    var titlematch: RegExpMatchArray | null = value.match(regexAgendaTitle);
    var durmatch: RegExpMatchArray | null = value.match(regexAgendaTime);

    if (titlematch === null) {
      console.log("No title match in [" + value + "]");
    }
    if (titlematch![1] === null) {
      console.log("No title capture group match in [" + value + "]");
    }
    if (durmatch === null) {
      console.log("No duration match in [" + value + "]");
      return;
    }

    var item: agendaItem = {
      duration: durationStringToSeconds(durmatch[0]),
      title: titlematch![1],
    };

    agendaItems.push(item); // returns number of items in array

    /* // DEBUG
    var count: number = agendaItems.push(item);
    console.log("Pushed item, returned: " + count!);
    console.log("parseAgendaItems found title:" + titlematch![1]);
    console.log("  for duration:" + durmatch[0]);
    console.log("  sec:" + durationStringToSeconds(durmatch[0]));
    //*/
  });

  return agendaItems.length;
}

/** updatePanel(): update all the panel data */
/**
 * @param timeobj: number; now; used for updating clock
 */
function updatePanel(timeStamp: number) {
  //eslint-disable-next-line
  if (!(itemCountLabel && itemTitleLabel)) {
    console.log("UI element unitialised in updatePanel()");
    return;
  }

  // update wall clock reading
  var timeobj: Date = new Date(timeStamp);
  updateClock(timeobj);

  // if done with agenda, don't update timer/title/etc.
  if (currentAgendaItemIndex >= agendaItems.length) return;

  // set title & counter
  var title: string | undefined = agendaItems[currentAgendaItemIndex].title;
  if (title != undefined) itemTitleLabel.textContent = title;

  // set item counter text
  // eslint-disable-next-line
  itemCountLabel.textContent =
    "item " + (currentAgendaItemIndex+1) + " of " + agendaItems.length;

  // set item remaining time text
  switch (runningFlag) {
    case 0:
      // eslint-disable-next-line
      if (currentAgendaItemIndex < agendaItems.length)
        updateTimer(agendaItems[currentAgendaItemIndex].duration);
      break;
    case 1:
      updateTimer(remainingTime(timeStamp));
      break;
    case 2: // don't update timer output during pause
    default:
      break;
  }

  return;
}

// CALLBACKS AND HANDLERS /////////////////

/** callbackClockTick(): main update callback; calls updatePanel() at end. */
function callbackClockTick() {
  // one shared timestamp across clocks
  var timeStamp = Date.now();

  // if we haven't advanced a second yet ... return.
  var nowSec: number = timeStamp / 1000; // only do the math once.
  if (nowSec === curTimeStamp) return; // no change
  curTimeStamp = nowSec; // remember for next run.

  // if timer is running, check for timeout.
  if (runningFlag === 1) {
    // if timeout, move to next item
    if (remainingTime(timeStamp) <= 0) nextItem();
  } // end 'timer is running'

  updatePanel(timeStamp); // do the actual widget updates

  return;
}

/** callbackStartPauseButton(): click handler*/
function callbackStartPauseButton(): void {
  if (!(startButton && pauseButton && timerOutput && startButtonLabel)) {
    console.log("HTML element was null in callbackStartPauseButton()");
    return;
  }

  switch (runningFlag) {
    case 0:
      // was stopped, pressed play to start
      startTime = Date.now();
      timerDuration = agendaItems[currentAgendaItemIndex].duration!;
      runningFlag = 1;
      pauseButton.style.display = "inline-block";
      startButton.style.display = "none";
      break;
    case 1:
    default:
      // was playing, pressed pause
      runningFlag = 2; // pause timer calback
      pauseButton.style.display = "none";
      startButton.style.display = "inline-block";
      startButtonLabel.textContent = "continue";
      pauseTimeStamp = Date.now();
      break;
    case 2:
      // was paused, pressed play to unpause
      // update timer duration to account for the dead/pause time
      timerDuration += Math.floor((Date.now() - pauseTimeStamp!) / 1000);
      startButton.style.display = "none";
      pauseButton.style.display = "inline-block";
      runningFlag = 1; // starts timer
      break;
  } // esac
}

/** callbackReloadAgenda(): search for agenda items */
export async function callbackReloadAgenda() {
  return Word.run(async (context) => {
    runningFlag = 0; // restart
    currentAgendaItemIndex = 0;

    // eslint-disable-next-line
    if (!(pauseButton && startButton &&
      startButtonLabel && nextButton && timerOutput)) {
      console.log("Element undefined in callbackReloadAgenda()");
      return;
    }

    pauseButton.style.display = "none";
    startButton.style.display = "inline-block";
    startButtonLabel.textContent = "start";
    // jic nextButton was hidden at end of meeting, and user is reloading
    nextButton.style.display = "inline-block";
    timerOutput.textContent = "-";

    var body: Word.Body = context.document.body;
    if (body === undefined) {
      console.log("Error loading document body in callbackReloadAgenda()");
      return;
    }

    body.load("text");
    await context.sync();
    const text = body.text.replace(/\r/g, "\n\r"); // for regex
    if (regexAgendaItem.test(text)) parseAgendaItems(text);
  }); // end Word.run()
} // end callbackReloadAgenda()

/** callbackNextButton(): jump to next Agenda Item */
function callbackNextButton() {
  nextItem();
  return;
}

/** callbackPreviousButton(): jump to previous Agenda Item */
function callbackPreviousButton() {
  if (currentAgendaItemIndex <= 0) return; // out of range
  currentAgendaItemIndex--; // move to prev item
  timerDuration = 60 * 5; // prep timer w/ 5min on clock
  startTime = Date.now();

  if (currentAgendaItemIndex + 1 == agendaItems.length) {
    // KISS solution taken. End of meeting sets runningFlag = 0.
    // When the flag is 0, updatePanel() sets timerDuration using
    // agendaItem[currentAgendaItem].duration. Right now, we don't
    // want that. We could set another flag and parse that in
    // updatePanel(), then wait for the user to explicitly press start.
    // That's not simple, though. This is easier and maybe less
    // bug-prone. If they've ended the meeting and click previous,
    // they must want to start the meeting again, right? So
    // we set runningFlag = 1.
    if (nextButton != null) nextButton.style.display = "inline-block";
    if (pauseButton != null) pauseButton.style.display = "inline-block";
    runningFlag = 1;
  }

  return;
}

/** callbackChangeDuration(): add/remove minutes from current event duration */
function callbackChangeDuration(this: any) {
  switch ((<Element>this).id) {
    case "addOneMinuteButton":
      changeDuration(1 * 60);
      break;
    case "addFiveMinutesButton":
      changeDuration(5 * 60);
      break;
    case "subtractOneMinuteButton":
      changeDuration(-1 * 60);
      break;
    case "subtractFiveMinutesButton":
      changeDuration(-5 * 60);
      break;
    default:
      break;
  }

  return;
}

// PROGRAM FLOW /////////////////

// main()-like thing
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // check that we have all of our UI elements ...
    /* eslint-disable */
    if (!(startButton && pauseButton && timerOutput
      && clockOutput && reloadButton && startButtonLabel
      && nextButton && previousButton && addOneMinuteButton
      && addFiveMinutesButton && subtractOneMinuteButton
      && subtractFiveMinutesButton)) {
      console.log("Error loading ui element(s) from DOM.");
      if (startButton === null) console.log("  startButton obj is null");
      if (startButtonLabel === null) console.log("  startButtonLabel obj is null");
      if (pauseButton === null) console.log("  pauseButton obj is null");
      if (previousButton === null) console.log("  previousButton obj is null");
      if (nextButton === null) console.log("  nextButton obj is null");
      if (timerOutput === null) console.log("  timerOutput obj is null");
      if (clockOutput === null) console.log("  clockOutput obj is null");
      if (reloadButton === null) console.log("  reloadButton obj is null");
      if (addFiveMinutesButton === null) console.log("  addFiveMinutesButton obj is null");
      if (addOneMinuteButton === null) console.log("  addOneMinuteButton obj is null");
      if (subtractFiveMinutesButton === null) console.log("  subtractFiveMinutesButton obj is null");
      if (subtractOneMinuteButton === null) console.log("  subtractOneMinuteButton obj is null");
      return;
    }
    /* eslint-enable */

    // attach button callbacks
    /* eslint-disable */
    reloadButton.addEventListener("click", callbackReloadAgenda);
    startButton.addEventListener("click", callbackStartPauseButton);
    pauseButton.addEventListener("click", callbackStartPauseButton);
    nextButton.addEventListener("click", callbackNextButton);
    previousButton.addEventListener("click", callbackPreviousButton);
    addOneMinuteButton.addEventListener("click", callbackChangeDuration);
    subtractOneMinuteButton.addEventListener("click", callbackChangeDuration);
    addFiveMinutesButton.addEventListener("click", callbackChangeDuration);
    subtractFiveMinutesButton.addEventListener("click", callbackChangeDuration);
    /* eslint-enable */

    setInterval(callbackClockTick, 100); // set up clock ticker
    callbackReloadAgenda(); // search doc for agenda items, update panel

    // display/un-hide our panel
    var doc = document.getElementById("app-body");
    if (doc != null) doc.style.display = "flex";
    else console.log("Word returned null doc object in Office.onReady().");
  } // end if (this is a valid Word application)
});
