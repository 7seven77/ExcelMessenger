# NellyDilemmaIM

In the music video for the song [Dilemma by Nelly (Feat. Kelly Rowland)](https://www.youtube.com/watch?v=8WYHDfJDPDc), Kelly Rowland can be seen sending a message using Excel.
I saw people ridiculing this and claiming this was stupid because it cannot be done.

I am planning to create some subroutines in vba as well as some pages for a website which allow Excel to be used as an "instant messaging" application

## Usage

- Create a website using the pages inside the `website` directory
  - Use the sql file to create the needed databases
    - Details about the database will need to be assigned in the `database.sql` file
- Create an excel spreadsheet with macros enabled
  - Use the `.vb` file to create the macros needed
  - Change the url to the one you are using
  - Change the cell locations from which data is retrieved to match what you need
  - Assign a 'button' or shortcut to send messages

These steps should work. If they don't, let me know and I will update them. :)

## Enable these under Tools -> References

- Microsoft XML, v6.0
- Microsoft HTML Object Library
