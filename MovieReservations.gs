function addMovieResevationsToCalendar() {
  var types = ['TOHO', 'Piccadilly', 'Sunshine'];
  for (index in types) {
    handleMails(types[index]);
  }
}

function handleMails(type) {
  var tytle;
  switch (type) {
    // TOHOシネマズ
    case 'TOHO':
      tytle = 'from:(i-net.ticket@ml.tohotheater.jp) subject:(チケット購入完了) is:unread';
      break;
    // 新宿ピカデリー
    case 'Piccadilly':
      tytle = '';
      break;
    // グランドサンシャイン池袋
    case 'Sunshine':
      tytle = 'from:(noreply@ticket-cinemasunshine.com) subject:(チケット予約完了) is:unread';
      break;
  }
  var start = 0;
  var max = 2;
 
  var threads = GmailApp.search(tytle,start,max);
  var messages = GmailApp.getMessagesForThreads(threads);
  
  for(let i=start; i<messages.length; i++){
    threads[i].markRead();
    for (message in messages[i]) {
      const body = messages[i][message].getPlainBody();
      switch (type) {
        // TOHOシネマズ
        case 'TOHO':
          handleTohoReservation(body);
          break;
        // 新宿ピカデリー
        case 'Piccadilly':
          handlePiccadillyReservation(body);
          break;
        // グランドサンシャイン池袋
        case 'Sunshine':
          handleSunshineReservation(body);
          break;
      }
    }
  }
};

function handleTohoReservation(body) {
  var resNum = body.split('■購入番号')[1].split('■電話番号').replace(/\s/g, "");
  var date = body.split('■上映日')[1].split('■時間').replace(/\s/g, "");
  var time = body.split('■時間')[1].split('■スクリーン').replace(/\s/g, "");
  var place = body.split('■映画館')[1].split('■作品名').replace(/\s/g, "");
  var title = body.split('■作品名')[1].split('■上映日').replace(/\s/g, "");

  var startTime = time.split('〜')[0];
  var endTime = time.split('〜')[1];
  var startDateTime = date + ' ' + startTime;
  var endDateTime = date + ' ' + endTime;

  var resTitle = '映画「' + title + '」';
  var desc = '予約番号：' + resNum;

  addToCalendar(resTitle, startDateTime, endDateTime, {description: desc, location: place})
}

function handlePiccadillyReservation(body) {

}

function handleSunshineReservation(body) {
  var resNum = body.split('[予約番号]')[1].split('[鑑賞日時]')[0].replace(/\s/g, "");
  var dateTime = body.split('[鑑賞日時]')[1].split('[作品名]')[0].replace(/\s/g, "");
  var title = body.split('[作品名]')[1].split('[スクリーン名]')[0].replace(/\s/g, "");

  var year = parseInt(dateTime.split('年')[0]);
  var month = parseInt(dateTime.split('年')[1].split('月')[0])-1; // 月は0から
  var day = parseInt(dateTime.split('月')[1].split('日')[0]);
  var startTime = dateTime.split(')')[1].split('-')[0];
  var startHour = parseInt(startTime.split(':')[0]);
  var startMin = parseInt(startTime.split(':')[1]);
  var endTime = dateTime.split('-')[1];
  var endHour = parseInt(endTime.split(':')[0]);
  var endMin = parseInt(endTime.split(':')[1]);
  var startDateTime = new Date(year, month, day, startHour, startMin);
  var endDateTime = new Date(year, month, day, endHour, endMin);

  var resTitle = '映画「' + title + '」';
  var description = '予約番号：' + resNum;
  var place = 'グランドシネマサンシャイン 池袋, 日本、〒170-0013 東京都豊島区東池袋１丁目３０−３ グランドスケープ池袋 4～13F';

  addToCalendar(resTitle, startDateTime, endDateTime, description, place);
}

function addToCalendar(title, startDateTime, endDateTime, desc, place) {
  var calender = CalendarApp.getCalendarById('zhushejia0817@gmail.com');
  calender.createEvent(title, startDateTime, endDateTime, {description: desc, location: place});
}
