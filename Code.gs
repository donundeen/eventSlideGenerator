var futuredays = 14; // how far in the future to look for creating new events
var futuretime = 1000 * 60 * 60 * 24 * futuredays;
var backgroundColor = "#000000";
var textColor= "#ffffff";


var months = [
  "Jan","Feb","Mar", "Apr", "May", "June","July","Aug","Sep","Oct","Nov", "Dec"
];
var dows  = [
  "Sunday","Monday","Tuesday","Wednesday","Thursday","Friday", "Saturday"
];

function removeAllSlides(preso){
  var slides = preso.getSlides();
  for (var i = 1; i < slides.length; i++){
    slides[i].remove();
  }
}


function formatDateString(event){
  var start = new Date(event.start.dateTime);
  var end = new Date(event.end.dateTime);
  
  var startDay = dows[start.getDay()];
  var startDate = months[start.getMonth()] + " " +start.getDate();
  var startHours = start.getHours();
//  startHours++;
  var startAmPm = "am";
  if (startHours > 12){
    startHours = startHours - 12;
    startAmPm = "pm";
  }
    
  var endHours = end.getHours();
//  endHours++;
  var endAmPm = "am";
  if (endHours > 12){
    endHours = endHours - 12;
    endAmPm = "pm";
  }

  
  var endDay = dows[end.getDay()];
  var endDate = months[end.getMonth()] + " " +end.getDate();
  
  Logger.log(start);
  Logger.log(start.getDay());
  Logger.log(start.getMonth());
  Logger.log(end);
  
  var startMinutes = start.getMinutes();
  Logger.log(startMinutes);
  if(startMinutes < 10){
    startMinutes = "0"+ startMinutes; 
  }
  var endMinutes = end.getMinutes();
  if(endMinutes < 10){
    endMinutes = "0"+ endMinutes; 
  }  
  
  
  var timeString = startDay + ", " + startDate + " " + startHours + ":" + startMinutes + startAmPm + " - " + endHours + ":" + endMinutes + endAmPm;
  
  return timeString;
 
}


function cleanText(text){
  if(!text){
    return "";
  }
  text = text.replace(/<br\/?>/gi,"\n");
  text = text.replace(/<[^>]+>/g,"");
  text = text.replace(/&nbsp;/," ");
//  text = text.replace(/\n+/g,"\n");
  text = text.replace(/^\n/g,"");
  text = text.replace(/\n$/g,"");
//  Logger.log(text);
  return text;
  
}

function addSlideForEvent(preso, event){
  // see https://developers.google.com/apps-script/reference/slides/slide
//  Logger.log(JSON.stringify(event, null, " "));
  var title = event.summary;
  
  var pageWidth = preso.getPageWidth();
  var pageHeight = preso.getPageHeight();
  
  if(title.match(/Maker Hub Open Hours/)){
    return; 
  }
  
  var description = event.description;
  var defaultImageId = "1wrYxafK58m01NcG6u5SYByTBC2doBx6g";
  var timeString = formatDateString(event);
  
  
  var image = false;
  var imageId = defaultImageId;
  var defaultImage = true;
  if (event.attachments && event.attachments.length > 0){
    var attachment = event.attachments[0];
    var url = attachment.fileUrl;
    if (attachment.title.match(/^.*\.(jpg|jpeg|png|gif)$/i)){
      image = url;    
      imageId = attachment.fileId;
      defaultImage = false;
    }
  }
  Logger.log("**********************************");
  Logger.log("**********************************");
  Logger.log(title);
  Logger.log(description);
  Logger.log(image);
  Logger.log(imageId);
  Logger.log("**********************************");
  Logger.log("**********************************");
  
  var newSlide = preso.appendSlide();
  
  newSlide.getBackground().setSolidFill(backgroundColor);
  
  var titleLength = title.length;
  var charsPerLine = 23;
  var numLines = Math.ceil(titleLength / charsPerLine);
  var titleHeight = numLines * 42;
  
  var titleShape = newSlide.insertShape(SlidesApp.ShapeType.TEXT_BOX, pageWidth / 2, 10, pageWidth / 2, titleHeight);
  titleShape.getText().setText(title);
  titleShape.getText().getTextStyle().setBold(true).setFontSize(36).setForegroundColor(textColor).setFontFamily("Philosopher");
     
  var dateShape = newSlide.insertShape(SlidesApp.ShapeType.TEXT_BOX, pageWidth / 2, titleHeight+ 10, pageWidth / 2, 50);
  dateShape.getText().setText(timeString);
  dateShape.getText().getTextStyle().setBold(true).setFontFamily("Georgia").setFontSize(24).setForegroundColor(textColor);
  
  var dateHeight = dateShape.getHeight();
  
  if(description && description.trim() != ""){
    var descShape = newSlide.insertShape(SlidesApp.ShapeType.TEXT_BOX, pageWidth / 2, dateHeight + titleHeight, pageWidth / 2, 50);
    description = cleanText(description);
    descShape.getText().setText(description);
    descShape.getText().getTextStyle().setBold(false).setFontFamily("Georgia").setForegroundColor(textColor).setFontSize(22);
  }
  
  
  if(imageId){
    var imageBlob = DriveApp.getFileById(imageId);        
    var imageShape = newSlide.insertImage(imageBlob);
    var imageWidth = imageShape.getWidth();
    var imageHeight = imageShape.getHeight();
    var targetWidth = pageWidth / 2;
    var scaleFactor = targetWidth / imageWidth;
    
    var newWidth = scaleFactor * imageWidth;
    var newHeight = scaleFactor * imageHeight;
    
    imageShape.setWidth(newWidth);
    imageShape.setHeight(newHeight);
    
    imageShape.setTop((pageHeight - newHeight) / 2)
    imageShape.setLeft(0);
  }
  
  if(!defaultImage && defaultImageId != imageId){
    //1317 × 1043
    var dtargetWidth = 128;
    var dimageBlob = DriveApp.getFileById(defaultImageId);        
    var dimageShape = newSlide.insertImage(dimageBlob);
    var dimageWidth = dimageShape.getWidth();
    var dimageHeight = dimageShape.getHeight();
  //  var dtargetWidth = pageWidth / 2;
    var dscaleFactor = dtargetWidth / dimageWidth;
    
    var dnewWidth = dscaleFactor * dimageWidth;
    var dnewHeight = dscaleFactor * dimageHeight;
    
    dimageShape.setWidth(dnewWidth);
    dimageShape.setHeight(dnewHeight);
    
//    imageShape.setTop((pageHeight - newHeight) / 2)
    dimageShape.setTop((pageHeight - dnewHeight - 16) );

    dimageShape.setLeft(0);  
  }
  
  var locationShape = newSlide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 16, pageHeight - 48, pageWidth / 2, 50);
  locationShape.getText().setText("1st Floor Lau");
  locationShape.getText().getTextStyle().setBold(false).setFontFamily("Audiowide").setForegroundColor(textColor).setFontSize(32);
  
  
}

function getUpcomingEvents(startDate, endDate){
  var publicCal  = getPublicEventsCal();
  Logger.log(JSON.stringify(publicCal.getId(), null, " "));
//  var publicEvents = getCalEvents(publicCal, startDate, endDate);
  var publicEvents = getCalEventsAdvanced(publicCal.getId(), startDate, endDate);
  return publicEvents;
  
}



function run(){
    
  var startTime = new Date();
  var endTime = new Date(startTime.getTime() + futuretime);  
  
  var preso = getPublicEventsSlideshow();
  removeAllSlides(preso);
  
  // get events
  var events = getUpcomingEvents(startTime, endTime);
      
  for(var i=0; i< events.length; i++){
    var event = events[i];
    addSlideForEvent(preso, event);
  }
}



// utilities
function getCalEvents(cal, startTime, endTime){
  var realStartDate = new Date(startTime);
  var realEndDate = new Date(endTime);
  try{
    var events = cal.getEvents(new Date(startTime), new Date(endTime));
    var returnEvents = events.map(function(event){
      var returnEvent = {
        id: event.getId(),
        calendarName : cal.getName(),
        startTime: event.getStartTime().toString(),
        endTime : event.getEndTime().toString(),
        title : event.getTitle(),
        description : event.getDescription(),
        guestList : event.getGuestList(true).map(function(guest){
          return {
                  name : guest.getName(),
                  email : guest.getEmail(),
                  status: guest.getGuestStatus()
                 };
        }),
        creators : event.getCreators(),
        dateCreated : event.getDateCreated().toString(),
        location : event.getLocation(),
        color: event.getColor(),
        attachments : event.getAttachments()
      }
      return returnEvent;
    });
    return returnEvents;
  }catch (error){
    Logger.log(error);
    throw error; 
  }
}

function getCalEventsAdvanced(calId, startTime, endTime){
   var optionalArgs = {
    timeMin: startTime.toISOString(),
    timeMax: endTime.toISOString(),
    showDeleted: false,
    singleEvents: true,
    orderBy: 'startTime'
  };
  var response = Calendar.Events.list(calId, optionalArgs);
  var events = response.items;
  return events;
}