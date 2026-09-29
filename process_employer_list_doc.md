# Processing an employer BU list

Dartmouth sends us a bargaining-unit list each term. This document is the whole
process for getting that list into Broadstripes, from the email attachment
through to the announcement on Discord. Work through it top to bottom.

Plan on a couple of hours the first time, most of it waiting on Broadstripes
imports. The genuinely fiddly parts are step 2 (interpreting new program codes)
and step 5 (catching workers who changed their email address).

## Before you start

You need:

- **Broadstripes access** with admin rights — you'll be creating custom fields
  and running data imports.
- **Access to the Drive folder** [Employer BU
  Lists](https://drive.google.com/drive/u/2/folders/1P2venElfcMgoK_FFVQ5K4dslLVmg3OVs).
- **A working copy of [listwork-scripts](https://github.com/covertg/listwork-scripts)**
  with a python environment set up. See that repo's
  [README.md](https://github.com/covertg/listwork-scripts/blob/main/README.md)
  for setup. Steps 2 and 5 need a bit of comfort with the command line.
- **Someone to ask.** Julie / OLR can decode program codes we can't work out.

A note on privacy: these files have workers' home addresses and phone numbers in
them. Keep them in the Drive folder and in the repo's gitignored `data/`
directory; don't email them around.

---

## 1. Download the list and rename it with the date

Add a suffix with "YYYY.MM.DD" to the original filename, where the date is **the
date that Dartmouth sent it to us**. For example:

```
2026.07.28 TO GOLD Membership 26X.xlsx
```

That date is what identifies this list everywhere downstream — in the CSV
filename, in the Broadstripes custom field, and in the shared search. The
parsing script reads it straight out of the filename and will refuse to run
without it.

Save a copy of this renamed file to the Drive folder [Employer BU
Lists](https://drive.google.com/drive/u/2/folders/1P2venElfcMgoK_FFVQ5K4dslLVmg3OVs).

## 2. Reformat the file for Broadstripes

The main challenging part of this step is parsing the program/field of study
codes that Dartmouth gives us and converting them to the program and degree type
fields that we want to use in Broadstripes. It also entails some simpler tasks,
such as combining multiple address columns into one and separating a unified
name column into separate name columns. We do all of this with the script
`parse_employer_bu.py`.

1. Put the renamed `.xlsx` into the `data/` subdirectory of the repo.

2. **Look at the file first.** Dartmouth renames the columns fairly often, so
   check what you've got before parsing:

   ```bash
   python parse_employer_bu.py -i './data/2026.07.28 TO GOLD Membership 26X.xlsx' --show_columns
   ```

3. **Run the parse**, telling the script which columns to use. These arguments
   were correct for the 26W, 26S, and 26X lists:

   ```bash
   python parse_employer_bu.py \
       -i './data/2026.07.28 TO GOLD Membership 26X.xlsx' \
       --program_col "Program1 Code" \
       --fullname_col "LFM Name Formatted" \
       --address_cols "LO Street1" "LO Street2" "LO City" "LO State" "LO Zip"
   ```

   The repo's README has fuller usage notes.

4. **Quite possibly, this list will include new program/field of study codes that
   we haven't seen before.** The script will stop and list them rather than
   guess. When that happens:

   1. Manually deduce what the unknown codes represent, and add them to
      `program_mapping.toml`.
   2. It usually helps to look up the workers listed under that code — the
      script prints their Excel row numbers for you. Their `@dartmouth.edu`
      email often names their school (`.GR` Guarini, `.TH` Thayer, `.MED`
      Geisel, `.PH` public health), and the existing entries in
      `program_mapping.toml` show the naming patterns Dartmouth uses.
   3. Write to Julie/OLR if there are codes that you can't figure out on your
      own. Add a comment in the toml recording what you concluded and how
      confident you are — the next person will thank you.
   4. If a new code represents a department/program that we've never seen
      before, the organization will need to be created on Broadstripes' "Shops &
      depts" page. Codes can also change because a department was *renamed* — in
      26X, Dartmouth's `GREARS*` (Earth Sciences) codes became `GREAPS*` after
      the department became Earth and Planetary Sciences. In that case rename
      the existing Broadstripes org rather than creating a second one, so the
      department's history stays in one place.
   5. Once the program mappings are figured out, re-run the script with the same
      arguments as before.

5. **Read the warnings.** The script prints `!! WARNING` for rows that look
   wrong but doesn't stop for them, because some of what it flags is legitimate.
   It gives you Excel row numbers so you can look each one up in the `.xlsx`:

   - *Rows sharing a name.* Usually the same worker exported twice. Uploading
     both makes a mess; delete the bad copy.
   - *Shifted address columns.* Dartmouth's export sometimes pushes a row's
     city → state → zip → phone one column to the right. The visible symptom is
     an empty city with a city name sitting in the state column. **The important
     part is the far end: the worker's ZIP code lands in the phone column**, and
     Broadstripes will cheerfully import it as their phone number. Fix these by
     hand in the `.xlsx` (shift the values back one column and clear the phone
     cell, since the real phone number was pushed off the end of the row) and
     re-run. In the 26X list, 7 rows were affected.
   - *A US state next to a malformed zip.* A dropped leading zero (`3755`) or
     an extra digit (`037555`) — a typo in the worker's own record rather than
     an export bug. One or two turn up in most lists, and they go into
     Broadstripes verbatim if you don't fix them.
   - *States or zips that aren't US-shaped.* Usually just an international
     address. Eyeball and move on.

   If you edit the `.xlsx`, save the corrections as a *new copy* — keep the file
   Dartmouth sent us unmodified, and keep the date in the filename.

6. The script writes a CSV next to the input file, named like `BU List Employer
   2026.07.28 made 2026.09.04_22.48.42.csv`.

7. Scan through the CSV to make sure there are no obvious discrepancies. Check
   in particular that the `Employer` and `Degree` columns look sane and that the
   row count matches what you expect.

8. This output file is our parsed version of the employer BU.

## 3. Create a new "Custom Field" in Broadstripes

This is what marks the workers who are part of this term's bargaining unit.

Go to Settings → Custom Fields and create a new custom field named `BU List
Employer YYYY.MM.DD`, where the date is **the date that Dartmouth sent it to
us** — e.g. `BU List Employer 2026.07.28`. Use field type Checkbox.

**The name has to match exactly**, four-digit year and all. The parsing script
has already put a column with this exact name into the CSV, and Broadstripes
matches the two up by name; a mismatch means nobody gets checked off. Copy the
header out of the CSV if you want to be sure, and check it against the previous
term's field while you're in Settings.

---

We'll now upload the parsed employer BU (CSV file) to Broadstripes! To keep the
database clean, this happens in two upload steps for two different populations
of workers: workers that already exist in Broadstripes, and then workers that
are new to our database.

## 4. Update existing workers, and identify the (likely) new ones

1. In Broadstripes, go to Settings → Data imports. Hit "+ New".
2. Give the upload an optional name such as "YYYY Term employer BU".
3. Upload the parsed employer BU (CSV file).
4. **Define data mappings.** The exact names will vary based on the format
   Dartmouth gives us, but be sure to include mappings to the Broadstripes
   fields for: Nickname, First Name, Middle Name, Last Name, Personal Cell
   Phone, Business Email, Employer, Degree, Home Address. Everything else can be
   skipped.
   - Check "Match?" on Business Email.
   - This should look like the following:
     ![][image1]
5. **Configure the merge/append behavior** (still on the "Data mappings" tab):
   1. Keep default settings under "Matched records". This should be: use the
      data to update records; append new data to existing; update the existing
      employment.
   2. **Change the defaults for un-matched records** to skip them and make data
      available for download.
   3. Make sure that "Automatically create shops and departments and link
      employments" is checked.
   4. This should look like:
      ![][image2]
6. Hit Preview and let the upload do a dry-run. This takes a couple of minutes.
   Check the Errors and Skips to see if there are any issues to be corrected,
   and give its summary a general sanity-check.
7. Finally, hit Submit!
8. After a couple of minutes, Broadstripes will have updated the existing
   workers' information. It will also give you a list of "Skips", which
   represents all of the (likely) new employees. The last two steps of this
   process use these Skips.

## 5. Check the "Skips" for workers who just changed their email

When a worker changes their email address (think name changes due to marriage,
transition, etc.) they won't get "matched" by the first import, and they'll show
up in the Skips as though they were brand new. We want to catch these rather
than making duplicate entries.

1. Download the Skips file from Broadstripes and run `check_skipped_imports.py`
   against it. **TODO: how to find skips using the script.**

2. For each duplicate the script turns up:
   1. Look up the worker on Broadstripes by searching for their name and hitting
      enter. Use the CAT layout or some other layout that shows their primary
      contact info. Click on the email address that's listed for them (this will
      be their out-of-date address), and a pop-up window will appear that lets
      you modify their contact info.

      (You can also get here by navigating to the worker's page itself and
      clicking the "Edit" tab. But editing from the search view is a bit more
      convenient.)
   2. Input their new email address, saving it as a Business Email. Also click
      the greyed-out rectangle next to their newly-input email; it will become
      green and say Primary to indicate that this is now their primary email
      address.
   3. You don't need to delete their old email address. It might be useful to
      keep it around for record-keeping's sake.

   This is never going to be perfect, and that's fine — the goal is to catch the
   obvious ones, not to be exhaustive.

## 6. Upload the file again to create the new workers

1. Use the same data mappings as before.
2. Use the same merge/append behavior for matched records. **But for unmatched
   records**, this time **use the data to create new records**. Keep the
   Broadstripes defaults. (No need to "Avoid creating duplicates" — I've
   experienced that feature not always working reliably, which is why we do our
   own duplicate searching in step 5.)
3. Preview and submit as in the first upload.
4. Once done, a Broadstripes search for `BU List Employer YYYY.MM.DD` (with the
   date set appropriately) should pull up all the workers you've just updated
   and added. We're almost done!

## 7. Create a new Broadstripes shared search for this term's BU

## 8. (TODO) "Entered BU Term", or "FirstTermThisTerm"

## 9. Announce on Discord #listwork

---

## Todo later

- Mailchimp contacts synchronization SOPs in general, including keeping staff on
  the mail list.
