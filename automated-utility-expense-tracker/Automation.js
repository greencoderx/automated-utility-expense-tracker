function autoUpdateMonthlyBills() {
  Logger.log("Starting monthly bill automation...");
  
  // 1. TIME Internet
  extractTimePDFsToDrive();
  parseTimePDFsToSheet();
  
  // 2. SK Magic
  parseSKMagicToSheet();
  
  // 3. TNB Electric (Boost)
  parseTNBToSheet();
  
  Logger.log("All monthly bill automations completed successfully.");
}
