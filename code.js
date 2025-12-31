function sendChristmasInvites() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {

    var name = data[i][0];        // Attendee Name
    var email = data[i][1];       // Attendee Email
    var status = data[i][2];      // Yes / Sent

    if (status && status.toString().trim().toLowerCase() === "yes") {

      var subject = "Invitation for Christmas Party";

      var htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
</head>
<body>
  <div style="font-family: Arial, sans-serif; background: linear-gradient(180deg, #1b1b2f, #16213e); padding:20px;">
    <div style="max-width:600px; margin:auto; background:#ffffff; border-radius:10px; padding:25px; text-align:center;">

      <h1 style="color:#c0392b;">
        🎄 Merry Christmas ${name}!
      </h1>

      <p style="font-size:16px; color:#333;">
        This is an <b>informal friends-only Christmas get-together</b>.
        Just for fun, food, laughs & good vibes.
      </p>

      <hr style="margin:20px 0;">

      <h2 style="color:#27ae60;">Party Details</h2>

      <p style="font-size:16px; text-align:left; display:inline-block;">
        📅 <b>Date:</b> 22 December 2025<br><br>
        ⏰ <b>Time:</b> 7:30 PM onwards<br><br>
        📍 <b>Venue:</b> Event Venue, Example City<br><br>
        👕 <b>Dress Code:</b> White top (bottom of your choice)
      </p>

      <div style="margin:25px 0; padding:15px; background:#f1f8e9; border-radius:8px;">
        <p style="font-size:15px; color:#2c3e50;">
          🎟️ <b>This email is your entry pass</b>
        </p>
      </div>

      <h2 style="color:#c0392b; margin-top:30px;">Organizers</h2>

      <div style="text-align:center;">

        <div style="display:inline-block; width:150px; margin:20px;">
          <img src="https://via.placeholder.com/110" style="border-radius:50%;">
          <p style="font-weight:bold;">Organizer One</p>
          <p style="font-size:13px;">organizer1@example.com</p>
        </div>

        <div style="display:inline-block; width:150px; margin:20px;">
          <img src="https://via.placeholder.com/110" style="border-radius:50%;">
          <p style="font-weight:bold;">Organizer Two</p>
          <p style="font-size:13px;">organizer2@example.com</p>
        </div>

        <div style="display:inline-block; width:150px; margin:20px;">
          <img src="https://via.placeholder.com/110" style="border-radius:50%;">
          <p style="font-weight:bold;">Organizer Three</p>
          <p style="font-size:13px;">organizer3@example.com</p>
        </div>

        <div style="display:inline-block; width:150px; margin:20px;">
          <img src="https://via.placeholder.com/110" style="border-radius:50%;">
          <p style="font-weight:bold;">Organizer Four</p>
          <p style="font-size:13px;">organizer4@example.com</p>
        </div>

      </div>

      <p style="font-size:14px; color:#555;">
        Dress well | Bring smiles | Come hungry
      </p>

      <div style="margin:30px 0; text-align:center;">

        <a href="https://calendar.google.com/calendar/render?action=TEMPLATE&text=Christmas+Party&details=Friends+Christmas+Get-together&location=Event+Venue&dates=20251222T140000Z/20251222T182900Z"
           style="display:inline-block; padding:12px 22px; margin:10px;
                  background:#d32f2f; color:white; text-decoration:none;
                  border-radius:6px; font-size:15px; font-weight:bold;">
          Add to Calendar
        </a>

        <a href="https://www.google.com/maps/search/?api=1&query=Event+Venue"
           style="display:inline-block; padding:12px 22px; margin:10px;
                  background:#1976d2; color:white; text-decoration:none;
                  border-radius:6px; font-size:15px; font-weight:bold;">
          View Location
        </a>

      </div>

      <p style="margin-top:30px; font-size:13px; color:#999;">
        See you soon!<br>
        – Event Team –
      </p>

    </div>
  </div>
</body>
</html>
      `;

      GmailApp.sendEmail(
        email,
        subject,
        "You are invited to a Christmas party.",
        { htmlBody: htmlBody }
      );

      // Mark as sent
      sheet.getRange(i + 1, 3).setValue("Sent");
    }
  }
}
