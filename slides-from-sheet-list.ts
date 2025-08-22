/**
 * Auto-populate Google Slides with names from Google Sheets
 * 
 * Setup Instructions:
 * 1. Open Google Apps Script (script.google.com)
 * 2. Create a new project and paste this code
 * 3. Update the SPREADSHEET_ID and PRESENTATION_ID variables below
 * 4. Update the SHEET_NAME and name column reference if needed
 * 5. Run the script
 */

// Configuration - UPDATE THESE VALUES
const SPREADSHEET_ID = 'your-spreadsheet-id-here'; // Get from the URL of your Google Sheet
const PRESENTATION_ID = 'your-presentation-id-here'; // Get from the URL of your Google Slides
const SHEET_NAME = 'Sheet1'; // Name of your sheet tab
const NAME_COLUMN = 'A'; // Column containing the names (A, B, C, etc.)
const TEMPLATE_SLIDE_INDEX = 1; // Index of your template slide (0-based, so slide 2 = index 1)

function populateSlidesFromSheet() {
  try {
    // Open the spreadsheet and presentation
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const presentation = SlidesApp.openById(PRESENTATION_ID);
    
    // Get the sheet and names
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    const names = getNamesFromSheet(sheet);
    
    if (names.length === 0) {
      console.log('No names found in the spreadsheet');
      return;
    }
    
    // Format names to show first name and last initial
    const formattedNames = names.map(name => formatName(name));
    
    // Randomize the names order
    const randomizedNames = shuffleArray(formattedNames);
    console.log('Names have been formatted and randomized');
    
    // Get the template slide
    const slides = presentation.getSlides();
    const templateSlide = slides[TEMPLATE_SLIDE_INDEX];
    
    if (!templateSlide) {
      throw new Error(`Template slide not found at index ${TEMPLATE_SLIDE_INDEX}`);
    }
    
    console.log(`Found ${randomizedNames.length} names. Creating slides in randomized order...`);
    
    // Create a slide for each name (now in randomized order)
    randomizedNames.forEach((name, index) => {
      createSlideForName(presentation, templateSlide, name, index + 1);
    });
    
    console.log(`Successfully created ${randomizedNames.length} slides in randomized order!`);
    
  } catch (error) {
    console.error('Error:', error.toString());
  }
}

function getNamesFromSheet(sheet) {
  // Get all data from the name column
  const range = sheet.getRange(NAME_COLUMN + ':' + NAME_COLUMN);
  const values = range.getValues();
  
  // Filter out empty cells and header (assuming first row might be header)
  const names = values
    .map(row => row[0])
    .filter(name => name && name.toString().trim() !== '')
    .slice(1); // Remove first row (header)
  
  return names;
}

// Function to randomize/shuffle an array using Fisher-Yates algorithm
function shuffleArray(array) {
  // Create a copy of the array to avoid modifying the original
  const shuffled = [...array];
  
  // Fisher-Yates shuffle algorithm
  for (let i = shuffled.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
  }
  
  return shuffled;
}

// Function to format names as "FirstName L."
function formatName(fullName) {
  if (!fullName || typeof fullName !== 'string') {
    return fullName; // Return as-is if not a valid string
  }
  
  // Clean up the name (remove extra spaces, trim)
  const cleanName = fullName.toString().trim().replace(/\s+/g, ' ');
  
  // Split the name into parts
  const nameParts = cleanName.split(' ');
  
  if (nameParts.length === 1) {
    // Only one name provided, return as-is
    return nameParts[0];
  }
  
  // Get first name and last initial
  const firstName = nameParts[0];
  const lastName = nameParts[nameParts.length - 1]; // Get the last part as surname
  const lastInitial = lastName.charAt(0).toUpperCase();
  
  return `${firstName} ${lastInitial}.`;
}

function createSlideForName(presentation, templateSlide, name, slideNumber) {
  // Duplicate the template slide
  const newSlide = templateSlide.duplicate();
  
  // Move the new slide to the end
  const slides = presentation.getSlides();
  const lastPosition = slides.length - 1;
  newSlide.move(lastPosition);
  
  // Replace placeholder text with the name
  replacePlaceholderText(newSlide, name);
  
  console.log(`Created slide ${slideNumber} for: ${name}`);
}

function replacePlaceholderText(slide, name) {
  // Get all text elements on the slide
  const textElements = slide.getPageElements()
    .filter(element => element.getPageElementType() === SlidesApp.PageElementType.SHAPE)
    .map(shape => shape.asShape().getText());
  
  // Replace placeholders in all text elements
  textElements.forEach(textRange => {
    // Replace common placeholders - customize these as needed
    textRange.replaceAllText('{{NAME}}', name);
    textRange.replaceAllText('{{name}}', name);
    textRange.replaceAllText('[NAME]', name);
    textRange.replaceAllText('[name]', name);
    textRange.replaceAllText('NAME_PLACEHOLDER', name);
    
    // You can add more placeholder patterns here
  });
}

// Optional: Function to clean up - removes all slides except title and template
function cleanupSlides() {
  const presentation = SlidesApp.openById(PRESENTATION_ID);
  const slides = presentation.getSlides();
  
  // Keep only the first slide (title) and template slide
  const slidesToKeep = 2; // Adjust this number based on your setup
  
  for (let i = slides.length - 1; i >= slidesToKeep; i--) {
    slides[i].remove();
  }
  
  console.log(`Cleanup complete. Kept ${slidesToKeep} slides.`);
}

// Optional: Test function to verify your setup
function testSetup() {
  try {
    // Test spreadsheet access
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('✓ Spreadsheet access successful');
    
    // Test sheet access
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    console.log('✓ Sheet access successful');
    
    // Test presentation access
    const presentation = SlidesApp.openById(PRESENTATION_ID);
    console.log('✓ Presentation access successful');
    
    // Test template slide
    const slides = presentation.getSlides();
    const templateSlide = slides[TEMPLATE_SLIDE_INDEX];
    console.log(`✓ Template slide found: "${templateSlide.getObjectId()}"`);
    
    // Test names retrieval and formatting
    const names = getNamesFromSheet(sheet);
    const formattedNames = names.map(name => formatName(name));
    console.log(`✓ Found ${names.length} names`);
    console.log('Original names:', names.slice(0, 3));
    console.log('Formatted names:', formattedNames.slice(0, 3));
    
  } catch (error) {
    console.error('❌ Setup test failed:', error.toString());
  }
}