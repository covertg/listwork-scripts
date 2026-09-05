1. **Download the employer BU and rename it with date**: add a suffix with “YYYY.MM.DD” to the original filename, where the date is **the date that Dartmouth sent it to us**. Save a copy of this file to the Drive folder [Employer BU Lists](https://drive.google.com/drive/u/2/folders/1P2venElfcMgoK_FFVQ5K4dslLVmg3OVs).

2. **Reformat this file in a way that is friendly for Broadstripes and our usage**. The main challenging part of this step is parsing the program/field of study codes that Dartmouth gives us and converting them to the program and degree type fields that we want to use in Broadstripes. This step also entails some other simpler tasks, such as combining multiple address columns into one column, and separating a unified name column (given as “Last First Middle”) into separate name columns. Currently, we do this reformatting automatically via a custom python script. Using it requires a bit of comfort with the command line and using python:  
   1. If you have not yet, git clone or otherwise download the repository [listwork-scripts](https://github.com/covertg/listwork-scripts).  
   2. Place the renamed employer BU xlsx file into the “data” subdirectory of this repo.  
   3. Use the script `parse_employer_bu.py` to parse the xlsx file. Refer to the README.md documentation file for details regarding usage of the script.  
   4. Quite possibly, this BU list will include new program/field of study codes that we haven’t seen before. The script will raise an error if so…  
      1. In which case, manually deduce what the unknown codes represent, and update `program_mapping.toml` accordingly.  
      2. It usually helps to look up the workers that are named under that degree code. Write to Julie/OLR if there are codes that you can’t figure out on your own.  
      3. If a new code represents a department/program that we’ve never seen before, the organization will need to be created on Broadstripes’ “Shops & depts” page.  
      4. Once the program mappings are figured out, re-run the script with the same arguments as before.  
   5. The script will create and save a CSV file.  
   6. Scan through the CSV file to make sure there are no obvious discrepancies.  
   7. This output file is our parsed version of the employer BU.

   

3. **Create a new “Custom Field”** in Broadstripes to represent the workers that are a part of this employer BU: go to Settings → Custom Fields. Create a new custom field named “BU List Employer YY.MM.DD”, where the date is **the date that Dartmouth sent it to us**. Use field type Checkbox. Make sure it follows the same naming format as the previous BU List custom fields.

   

We’ll now upload the parsed employer BU (CSV file) to Broadstripes\! To try to keep the database clean, this actually happens in two upload steps for two different populations of workers: workers that already exist in Broadstripes, and then workers that are new to our database.

4. **Update existing workers’ data and identify the (likely) new workers.**  
   1. In Broadstripes, go to Settings → Data imports. Hit “+ New”.  
   2. You can give this upload an optional name such as “YYYY Term employer BU”  
   3. Upload the parsed employer BU (CSV file).  
   4. Define data mappings:  
      1. The exact names will vary based on the format that Dartmouth gives us, but be sure to include mappings to the Broadstripes fields for: Nickname, First Name, Middle Name, Last Name, Personal Cell Phone, Business Email, Employer, Degree, Home Address. Everything else can be skipped.  
      2. Check “Match?” on Business Email.  
      3. This should look like the following:  
         ![][image1]  
   5. Configure the merge/append behavior (still on “Data mappings tab”):  
      1. Keep default settings under “Matched records”. This should be: use the data to update records; append new data to existing; update the existing employment.  
      2. **Change the defaults for un-matched records** to skip them and make data available for download.  
      3. Make sure that “Automatically create shops and departments and link employments” is checked.  
      4. This should look like:  
         ![][image2]  
   6. Hit Preview and let the upload do a dry-run. (This will take a couple of minutes). Check the Errors and Skips to see if there are any issues to be corrected. Also give a general sanity-check of its summary.  
   7. Finally, hit Submit\!  
   8. After a couple of minutes, Broadstripes will have updated the existing workers’ information. It will also give you a list of “Skips”, which represents all of the (likely) new employees. The last two overall steps of this process will use these Skips.

5. **Look through the list of “Skips” to make sure that all workers are actually new, and update worker email addresses if needed to avoid duplicates.** When a worker changes their email address (think name changes due to marriage, transition, etc.) they might not get “matched” by the first import. We want to try to catch these rather than making duplicate entries.  
   1. TODO: how to find skips using the script.  
   2. For each duplicate:  
      1. Look up the worker on Broadstripes by searching for their name and hitting enter. Use the CAT layout or some other layout that shows their primary contact info. Click on the email address that’s listed for them (this will be their out-of-date email address), and a pop-up window will appear that lets you modify their contact info.  
         (You can also get here by navigating to the worker’s page itself, and clicking on the “Edit” tab. But I find that editing from the search view is just a bit more convenient.)  
      2. Input their new email address, saving it as a Business Email. Also click the greyed-out rectangle next to their newly-input email; it will become green and say Primary to indicate that this is now their primary email address.  
      3. You don’t need to delete their old email address. It might be useful to keep it around for record-keeping sake.  
      4. This should look like:  
         

   

6. **Upload the repeats file to Broadstripes**.  
   1. Use the same data mappings as before.  
   2. Use the same merge/append behavior for matched records. **But for unmatched records**, this time **Use the data to create new records**. Keep the broadstripes defaults. (No need to “Avoid creating duplicates.” I’ve experienced this feature not always working reliably. Which is why we do our own automated duplicate searching anyways…)  
   3. Preview and submit as in the first upload\!  
   4. Once done, then a Broadstripes search for “BU List Employer YY.MM.DD” (with the date set appropriately) should pull up all the workers that you’ve just updated and added\! We’re almost done\!

   

7. **Create a new Breadstripes shared search for this term’s BU.**  
8. **(Todo) “Entered BU Term”, or “FirstTermThisTerm”**  
9. **Announce on Discord \#listwork**

Todo later:  
Mailchimp contacts synchronization SOPs in general, including keep staff on the mail list
