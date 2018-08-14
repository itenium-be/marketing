Email signature
===============

![Signature example](example_signature.PNG)

- In gmail, goto Settings > General > Signature
- Edit `signature.html`
- Open signature.html in a browser
- Select all and copy, then paste in your email settings
- Save Changes
- Reload page
- **ATTN:** Now manually add `text-decoration:none !important; text-decoration:none;` AGAIN to the `<a>`s with Inspect Elements
- Make a (silly) change to the signature to re-enable the "Save Changes" button
- Save Changes


Workaroundish stuff
-------------------

`text-decoration:none !important; text-decoration:none;`
-> For Office365

The manual Inspect Elements `text-decoration` changes are required to remove the underline in Outlook.com


Links
-----

**CSS Support in Email clients**:

https://www.campaignmonitor.com/css
