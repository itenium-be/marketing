Email signature
===============

![Signature example](example_signature.PNG)

Gmail:
===

- In gmail, goto Settings > General > Signature
- Edit `signature.html`
- Open signature.html in a browser
- Select all and copy, then paste in your email settings
- Save Changes
- Reload page
- **ATTN:** Now manually add `text-decoration:none !important; text-decoration:none;` AGAIN to the style attribute of all the `<a>` tags(email,mobile,website) with Inspect Elements 
- Make a (silly) change to the signature to re-enable the "Save Changes" button. Try a whitespace or a change in the name you later revert with another save changes.
- Save Changes

Explanation
-------------------

`text-decoration:none !important; text-decoration:none;`
-> For Office365

The manual Inspect Elements `text-decoration` changes are required to remove the underline 
-> For Outlook.com

Outlook 2016 Desktop App:
===

Links are underlined after copy pasting the signature. Select each link and press CTRL+U twice to remove.


Links
===

**CSS Support in Email clients**:

https://www.campaignmonitor.com/css
