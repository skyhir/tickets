/**
 * Combined onOpen function to add all custom menu items
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Citations Tools')
    .addItem('Add a Ticket', 'showAddTicketForm')
    .addItem('Batch Upload', 'showBatchTicketForm')
    .addItem('Process Payment', 'processTicketPayments')
    .addItem('Email Drivers', 'emailDriversTicketNotifications')
    .addItem('Add a Ticket (Blink)', 'showAddTicketFormBlink')
    .addSeparator()
    .addItem('Lookup Booking for Selected Row', 'manualBookingLookupForRow')
    .addSeparator()
    .addItem('Export Historical Tolls → CSV', 'exportHistoricalTollsCsv')
    .addToUi();
  
}