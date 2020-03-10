# usefulvbscript

VBS files that I have developed for work that do various useful things.

<table>
  <th colspan="2" align="left">SCRIPT Update File .xls to .xlsx</th>
  <tr>
    <td>SUMMARY</td><td>Changes the .xls files in a folder to .xlsx files.</td>
  </tr>
  <tr>
    <td>CHALLENGE</td>
    <td>Many companies have folders full of old format .xls files. When you need to use advanced Excel features, or copy and paste between new and old Excel files there are issues.</td>
  </tr>
  <tr>
    <td>INSTRUCTIONS</td>
          <td>
            1) Put this VBS file into a folder that has .xls files, run the script and click OK.
            <br>2) All .xls files will be opened one at a time, saved as a .xlsx file, then the .xls file is moved to the LegacyArchive folder.
            <br>3) At the end a log text file is created to show all changes made.
          </td></tr>
</table>

<table>
  <th colspan="2" align="left">SCRIPT Printer List</th>
  <tr>
    <td>SUMMARY</td>
    <td>Displays a list of all printers available on a computer.</td>
  </tr>
  <tr>
    <td>CHALLENGE</td>
    <td>I was creating an HTA that has buttons to run external vbscripts to automate printing of various reports. Some of those reports print in black and white and others must print in color. I created a dropdown list for the user to select their color printer since each computer listed our Xerox printer as something slightly different. I created this script to pull printer information from the computer.</td>
  </tr>
  <tr>
    <td>INSTRUCTIONS</td>
    <td>
      1) Put this VBS file anywhere on the computer, run the script and it will show a list of all available printers.
    </td></tr>
</table>
