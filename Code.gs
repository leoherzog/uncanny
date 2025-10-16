const cache = CacheService.getScriptCache();
const properties = PropertiesService.getScriptProperties();

function setIdeasSheetID() {
  properties.setProperty('sheetId', '18396IY1-OKj8jcWbwkDUOOmW0s7OhCJqfjOLx9xr7cg');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Uncanny Feedback')
    .setFaviconUrl('https://favicon.show/letter/U?color=3C71F7')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUser() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Unable to detect user email');
  return email;
}

function refreshCache_() {
  let sheet = SpreadsheetApp.openById(properties.getProperty('sheetId'));
  let ideas = sheet.getDataRange().getValues();
  ideas = JSON.stringify(ideas);
  cache.put('ideas', ideas, 60);
  return ideas;
}

function getIdeas(forceCacheRefresh = false) {
  let ideas = cache.get('ideas');
  if (!ideas || forceCacheRefresh) ideas = refreshCache_();
  console.log(ideas);
  ideas = JSON.parse(ideas);
  ideas = ideas.filter(x => !!x[ideas[0].indexOf('Visible')]);
  ideas = JSON.stringify(ideas);
  return ideas;
}

function getIdea(id) {
  let ideas = getIdeas();
  ideas = JSON.parse(ideas);
  return JSON.stringify(ideas.find(x => x[ideas[0].indexOf('ID')] === id));
}

function addIdea(title, description, submittor, status = 'New') {
  SpreadsheetApp.openById(properties.getProperty('sheetId')).appendRow([false, status, Utilities.getUuid(), title, description, new Date(), submittor, "[\"" + submittor + "\"]"]);
  refreshCache_();
  return true;
}

function toggleVote(id, email) {

  const lock = LockService.getScriptLock();

  try {
    // wait up to 10 seconds to acquire the lock
    if (!lock.tryLock(10000)) {
      throw new Error('System is busy. Please try again in a moment.');
    }

    const sheet = SpreadsheetApp.openById(properties.getProperty('sheetId')).getSheetByName('Ideas');
    const ideas = sheet.getDataRange().getValues();

    const rowIndex = ideas.findIndex(row => row[ideas[0].indexOf('ID')] === id);

    // check if idea exists before accessing it
    if (rowIndex === -1) throw new Error('Idea with id ' + id + ' not found');

    const idea = ideas[rowIndex];

    let votes = [];
    const votersData = idea[ideas[0].indexOf('Voters')];
    if (votersData) {
      try {
        votes = JSON.parse(votersData);
        if (!Array.isArray(votes)) votes = [];
        votes = votes.filter(x => !!x);
      } catch (e) {
        votes = [];
      }
    }

    const voteIndex = votes.indexOf(email);
    let action = '';
    if (voteIndex === -1) { // not in voters array yet
      votes.push(email);
      action = 'added';
    } else {
      votes.splice(voteIndex, 1);
      action = 'removed';
    }

    sheet.getRange(Number(rowIndex) + 1, Number(ideas[0].indexOf('Voters') + 1)).setValue(JSON.stringify(votes));

    SpreadsheetApp.flush();

    refreshCache_();

    return {"id": id, "title": idea[ideas[0].indexOf('Title')], "action": action, "newCount": votes.length};

  } finally {
    // always release the lock, even if an error occurred
    lock.releaseLock();
  }
}

function test() {
  // console.log(getUser());
  // getIdeas(true);
  // console.log(getIdea('9a3967e1-a1b3-494e-9816-c73808fbc70e'));
}