function sendChristmasInvites() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {

    var name = data[i][0];
    var email = data[i][1];
    var contributed = data[i][2];

    if (
      contributed &&
      contributed.toString().trim().toLowerCase() === "yes"
    ) {

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

      <!-- Header -->
      <h1 style="color:#c0392b;">
        <img src="https://twemoji.maxcdn.com/v/latest/72x72/1f384.png" width="22" style="vertical-align:middle;">
        Merry Christmas ${name}!
      </h1>

      <p style="font-size:16px; color:#333;">
        This is an <b>informal friends-only Christmas get-together</b>.
        Just for fun, food, laughs & good vibes.
      </p>

      <hr style="margin:20px 0;">

      <!-- Party Details -->
      <h2 style="color:#27ae60;">
        <img src="https://twemoji.maxcdn.com/v/latest/72x72/1f381.png" width="20" style="vertical-align:middle;">
        Party Details
      </h2>

      <p style="font-size:16px; text-align:left; display:inline-block;">
        <img src="https://twemoji.maxcdn.com/v/latest/72x72/1f4c5.png" width="18"> 
        <b>Date:</b> 22 December 2025<br><br>

        <img src="https://twemoji.maxcdn.com/v/latest/72x72/23f0.png" width="18"> 
        <b>Time:</b> 7:30 PM onwards<br><br>

        <img src="https://twemoji.maxcdn.com/v/latest/72x72/1f4cd.png" width="18"> 
        <b>Venue:</b> 483 Terrace, Sector 4, Vaishali, Ghaziabad, 201010<br><br>

        <img src="https://twemoji.maxcdn.com/v/latest/72x72/1f454.png" width="18"> 
        <b>Dress Code:</b> White top and lower can be anything
      </p>

      <!-- Entry Pass -->
      <div style="margin:25px 0; padding:15px; background:#f1f8e9; border-radius:8px;">
        <p style="font-size:15px; color:#2c3e50;">
          <img src="https://twemoji.maxcdn.com/v/latest/72x72/1f39f.png" width="18">
          <b>This email is your entry pass</b><br>
        </p>
      </div>

      <!-- ORGANIZERS SECTION -->
      <div style="margin-top:30px; text-align:center;">
        <h2 style="color:#c0392b; margin-bottom:20px;">Organizers</h2>

        <!-- Row 1 -->
        <div style="display:block; text-align:center; margin-bottom:20px;">
          
          <!-- Organizer 1 -->
          <div style="display:inline-block; width:150px; margin:20px; text-align:center;">
            <div style="width:110px; height:110px; margin:auto; overflow:hidden; border-radius:50%;">
              <img src="https://drive.google.com/uc?export=view&id=1CtApMvyFTIcKl_pE8BTHI7P6sSdNEN9f"
                  style="width:100%; height:100%; object-fit:cover;">
            </div>
            <p style="margin:8px 0 2px; font-weight:bold;">Parkhi Sharma</p>
            <p style="margin:0; font-size:13px;">parkhi.sharma2007@gmail.com</p>
          </div>

          <!-- Organizer 2 -->
          <div style="display:inline-block; width:150px; margin:20px; text-align:center;">
            <div style="width:110px; height:110px; margin:auto; overflow:hidden; border-radius:50%;">
              <img src="https://drive.google.com/uc?export=view&id=1Ga0jb7p-LNhaQ8p5wiVuFc9HSYLiGMUY"
                  style="width:100%; height:100%; object-fit:cover;">
            </div>
            <p style="margin:8px 0 2px; font-weight:bold;">Arpit Jain</p>
            <p style="margin:0; font-size:13px;">jain.arpit3112@gmail.com</p>
          </div>

        </div>

        <!-- Row 2 -->
        <div style="display:block; text-align:center;">

          <!-- Organizer 3 -->
          <div style="display:inline-block; width:150px; margin:20px; text-align:center;">
            <div style="width:110px; height:110px; margin:auto; overflow:hidden; border-radius:50%;">
              <img src="https://drive.google.com/uc?export=view&id=1Sjdd9OIn7i-q3mDRxImEEc6AvzuPrLYr"
                  style="width:100%; height:100%; object-fit:cover;">
            </div>
            <p style="margin:8px 0 2px; font-weight:bold;">Naina Chattree</p>
            <p style="margin:0; font-size:13px;">chattreenaina@gmail.com</p>
          </div>

          <!-- Organizer 4 -->
          <div style="display:inline-block; width:150px; margin:20px; text-align:center;">
            <div style="width:110px; height:110px; margin:auto; overflow:hidden; border-radius:50%;">
              <img src="https://drive.google.com/uc?export=view&id=1k-UnKV6XK4A8HsSDSynTuIKUZmxVhE44"
                  style="width:100%; height:100%; object-fit:cover;">
            </div>
            <p style="margin:8px 0 2px; font-weight:bold;">Bhagavath M S</p>
            <p style="margin:0; font-size:13px;">msbhagavath@gmail.com</p>
          </div>

        </div>

      </div>


      <p style="font-size:14px; color:#555;">
        Dress well | Bring smiles | Come hungry
      </p>

      <!-- ACTION BUTTONS -->
      <div style="margin:30px 0; text-align:center;">

        <!-- Add to Calendar Button -->
        <a href="https://calendar.google.com/calendar/render?action=TEMPLATE&text=Christmas+Party&details=Informal+friends+Christmas+get-together&location=483+Terrace%2C+Sector+4%2C+Vaishali%2C+Ghaziabad%2C+201010&dates=20251222T140000Z/20251222T182900Z"
          style="display:inline-block; padding:12px 22px; margin:10px;
                  background:#d32f2f; color:white; text-decoration:none;
                  border-radius:6px; font-size:15px; font-weight:bold;">Add to Calendar
        </a>

        <!-- View Location in Maps Button -->
        <a href="https://maps.app.goo.gl/jv5XgxpxvLsaUi7y9"
          style="display:inline-block; padding:12px 22px; margin:10px;
                  background:#1976d2; color:white; text-decoration:none;
                  border-radius:6px; font-size:15px; font-weight:bold;">
          View Location
        </a>
      </div>


      <p style="margin-top:30px; font-size:13px; color:#999;">
        See you soon!<br>
        – BasementDwellers –
      </p>
      
    </div>
  </div>
</body>
</html>
      `;

      GmailApp.sendEmail(
        email,
        subject,
        "You are invited to a Christmas party. Please view this email in HTML format.",
        { htmlBody: htmlBody }
      );

      sheet.getRange(i+1, 3).setValue("Sent");
    }
  }
}
