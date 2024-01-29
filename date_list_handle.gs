

function listDatesBetween(startDate, endDate) {
  var currentDate = new Date(startDate);
  endDate = new Date(endDate);
  
  var datesList = [];
  while (currentDate <= endDate) {
    console.log(currentDate);
    datesList.push(new Date(currentDate));
    currentDate.setDate(currentDate.getDate() + 1);
  }
  return datesList;
}
