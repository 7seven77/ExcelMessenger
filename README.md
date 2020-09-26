# ExcelMessenger

Send and recieve messages using Excel

## Inspiration

In the music video for the song [Dilemma by Nelly (Feat. Kelly Rowland)](https://www.youtube.com/watch?v=8WYHDfJDPDc), Kelly Rowland can be seen sending a message using Excel.
I saw people ridiculing this and claiming this was stupid because it cannot be done which inspired me to make it.

## Setting up

An excel sheet without macros is provided. This can be used or you can make your own. 

- Create a website using the pages inside the `website` directory
  - Use the sql file to create the needed databases
    - Details about the database will need to be assigned in the `database.sql` file
- Create an excel spreadsheet with macros enabled
  - Create a module for each of the `.vb` files
  - Functions in `data.vb` should be changed so they read from the cells you require
  - Website URL in `requests.vb` should be changed to lead to yours
  - `subs.vb` can be used to change the interface

`subs.vb` are subroutines that deal with the displaying of information. These will change a lot if you are making your own layout.

These steps should work. If they don't, let me know and I will update them. :)

## Enable these under Tools -> References

- Microsoft HTML Object Library

## Collaborators

- 7seven77
