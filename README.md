# Pit Sign Generator – README

This PowerPoint contains a **Pit Sign Generator** that automatically creates one pit sign per team using an Excel team list.

This file is intended to be used by event staff or volunteers. No coding knowledge is required.

---

## What You Need

* **PowerPoint (desktop version)** for Windows or Mac

  > PowerPoint Web will NOT work
* **Excel** with a team list
* The **macro‑enabled PowerPoint file** (`.pptm`)

---

## Excel File Requirements

Your Excel sheet must be formatted exactly like this:

| Column A | Column B         |
| -------- | ---------------- |
| 5393     | Vinegar Vikings  |
| 8114     | The Pro Tractors |

Important rules:

* Column **A** = Team Number
* Column **B** = Team Name
* No blank rows in the middle of the list
* The sheet with the team list must be the **active tab**
* Leave Excel **open** while running the generator

---

## PowerPoint Template Slide

Inside the PowerPoint there is **one template slide** that:

* Has the correct pit sign design
* Contains the text placeholders:

  * `{{NUM}}`
  * `{{NAME}}`
* Each placeholder appears **twice** on the slide

⚠️ Do not change the placeholder text.

---

## How to Run the Pit Sign Generator

### Step 1: Open Files

1. Open the Excel file with the team list
2. Open the PowerPoint **.pptm** file

---

### Step 2: Enable Macros

If prompted:

* Click **Enable Content** or **Enable Macros**

Macros are required for the generator to work.

---

### Step 3: Run the Generator

1. In PowerPoint, press:

   * **Windows:** `Alt + F11`
   * **Mac:** `Option + Fn + F11`
2. The VBA editor opens
3. Press **F5** or click **Run ▶**
4. Select:
   **`BuildPitSigns_TopToBottom`**

---

### Step 4: Answer the Prompts

You will be asked:

**Template slide number**

* Enter the slide number of the template slide (usually `1`)

**Excel start row**

* Enter `2` if row 1 has headers
* Enter `1` if there are no headers

---

### Step 5: Done

PowerPoint will:

* Duplicate the template slide
* Replace `{{NUM}}` and `{{NAME}}`
* Create one slide per team
* Insert slides in **top‑to‑bottom order**

A confirmation message will appear when complete.

---

## Saving & Sharing (Important)

When finished:

1. Go to **File → Save As**
2. Choose **PowerPoint Presentation (.pptx)**

This removes the macro and makes the file safe to:

* Email
* Share with schools or admins
* Send to printers

⚠️ Do NOT share the `.pptm` file externally.

---

## Common Issues

**Only one slide was created**

* Excel may have blank rows
* Wrong sheet is active

**Slides still show {{NUM}} or {{NAME}}**

* Macros were not enabled
* Template placeholders were edited

**Nothing happens**

* Excel is not open
* Using PowerPoint Web instead of desktop

---

## Support

If something breaks:

* Close both Excel and PowerPoint
* Reopen them
* Make sure the template slide is untouched
* Try again

This generator can be reused for any event by simply swapping the Excel team list.

---

**End of README**
