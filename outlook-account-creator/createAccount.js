const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

async function createOutlookAccount(userData) {
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();

  try {
    // Navigate to Outlook signup page
    await page.goto('https://signup.live.com/signup', { waitUntil: 'networkidle2' });

    // Fill in the email field
    await page.waitForSelector('input[name="MemberName"]');
    await page.type('input[name="MemberName"]', userData.email);

    // Click Next
    await page.click('input[type="submit"]');

    // Wait for password field and fill it
    await page.waitForSelector('input[name="Password"]', { visible: true });
    await page.type('input[name="Password"]', userData.password);

    // Click Next
    await page.click('input[type="submit"]');

    // Fill in first name and last name
    await page.waitForSelector('input[name="FirstName"]', { visible: true });
    await page.type('input[name="FirstName"]', userData.firstName);
    await page.type('input[name="LastName"]', userData.lastName);

    // Click Next
    await page.click('input[type="submit"]');

    // Fill in country/region and birthdate
    await page.waitForSelector('select[name="Country"]', { visible: true });
    await page.select('select[name="Country"]', userData.country);

    await page.type('input[name="BirthMonth"]', userData.birthMonth);
    await page.type('input[name="BirthDay"]', userData.birthDay);
    await page.type('input[name="BirthYear"]', userData.birthYear);

    // Click Next
    await page.click('input[type="submit"]');

    // Wait for captcha to be solved manually
    console.log(`Please solve the captcha manually in the opened browser window for ${userData.email}.`);
    await page.waitForFunction(
      () => !document.querySelector('input[name="ProofOfWork"]') && !document.querySelector('.captcha'),
      { timeout: 0 }
    );

    // After captcha solved, click Next or submit
    await page.click('input[type="submit"]');

    // Wait for possible IP blocking or rate limiting error
    try {
      await page.waitForSelector('.error-message', { timeout: 5000 });
      const errorMessage = await page.$eval('.error-message', el => el.textContent);
      if (errorMessage.toLowerCase().includes('too many accounts') || errorMessage.toLowerCase().includes('ip')) {
        console.log('Error detected: Too many accounts created from this IP. Please change your IP and try again.');
        await browser.close();
        return { success: false, email: userData.email, error: errorMessage };
      }
    } catch (e) {
      // No error message detected
    }

    console.log(`Account creation process completed for ${userData.email}.`);
    await browser.close();
    return { success: true, email: userData.email };
  } catch (error) {
    console.error('Error during account creation:', error);
    await browser.close();

function saveAccountsToExcel(accounts, filename = 'created_accounts.xlsx') {
  const worksheetData = accounts.map(acc => ({
    Email: acc.email,
    Status: acc.success ? 'Success' : 'Failed',
    Error: acc.error || ''
  }));

  const worksheet = XLSX.utils.json_to_sheet(worksheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Accounts');

  const filePath = path.resolve(__dirname, filename);
  XLSX.writeFile(workbook, filePath);
  console.log(`Saved account results to ${filePath}`);
}
