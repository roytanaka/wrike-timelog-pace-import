function fetchTimelog(start, end, sheetName) {
  // Increase end date by 1 day to capture timelog from previous day
  const d = new Date(end);
  d.setDate(d.getDate() + 1);
  const newEnd = d.toISOString().split('T')[0];

  clearExports(sheetName);

  const dates = `trackedDate=${JSON.stringify({
    start: start,
    end: newEnd,
  })}`;

  const timelogUrl = encodeURI(`/folders/IEABU4QJI7777777/timelogs?${dates}`); // IEABU4QJI7777777 Wrike root folder
  const timelogReq = wrikeGetReq(timelogUrl);
  const usersReq = wrikeGetReq('/contacts');

  const [timelogRes, usersRes] = UrlFetchApp.fetchAll([timelogReq, usersReq]);

  const timelogData = JSON.parse(timelogRes).data;
  const usersData = JSON.parse(usersRes).data;

  const groupFind =
    sheetName === 'Creative Export' ? 'Creative Studio' : 'Prepress';

  const groupData = usersData.filter((contact) => contact.type === 'Group');

  const { memberIds } = groupData.find(
    (group) => group.firstName === groupFind
  );

  const filteredTimelog = timelogData.filter((log) =>
    memberIds.includes(log.userId)
  );

  if (filteredTimelog.length === 0) {
    throw new Error('No timelogs found for dates selected');
  }
  buildExport(filteredTimelog, usersData, sheetName);
}

function buildExport(timelogData, usersData, sheetName) {
  const paceIds = getSheetData('Pace IDs');
  const activityCodes = getSheetData('Activity Codes');

  const logs = timelogData.map((log) => {
    const taskReq = wrikeGetReq(`/tasks/${log.taskId}`);
    const res = UrlFetchApp.fetch(taskReq.url, taskReq);
    const { data } = JSON.parse(res);
    const docketRe = /^S[^: ]+/;
    log.job = docketRe.test(data[0].title)
      ? docketRe.exec(data[0].title)[0]
      : data[0].title;

    log.email = usersData.find(
      (user) => user.id === log.userId
    ).profiles[0].email;

    if (log.categoryId) {
      const activityLookup = activityCodes.find((row) =>
        row.includes(log.categoryId)
      );
      log.timelogCat = activityLookup[2];
      log.comment = log.comment
        ? `${log.comment} [${activityLookup[1]}]`
        : activityLookup[1];
    } else {
      log.timelogCat = 20515; //Creative - Design
    }
    const employeeRow = paceIds.find((row) => row.includes(log.email));
    if (employeeRow) {
      log.employeeId = employeeRow[0];
    } else {
      log.employeeId = `email: ${log.email} not found in Pace IDs tab`;
    }

    return log;
  });

  const exportValues = logs.map((log) => [
    'I',
    log.job,
    '01',
    log.employeeId,
    log.timelogCat,
    log.hours,
    log.comment,
  ]);

  setExport(exportValues, sheetName);
}

function wrikeGetReq(url) {
  const TOKEN = PropertiesService.getScriptProperties().getProperty('token');
  const BASEURL = 'https://www.wrike.com/api/v4';
  const auth = { Authorization: `Bearer ${TOKEN}` };
  const req = {
    url: BASEURL + url,
    method: 'get',
    headers: auth,
  };
  return req;
}

function setExport(data, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  const range = sheet.getRange(`A2:G${data.length + 1}`);
  range.setValues(data);
}

function clearExports(sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = sheet.getDataRange().getLastRow();
  if (lastRow > 1) {
    const range = sheet.getRange(`A2:G${lastRow}`);
    range.clearContent();
  }
}
