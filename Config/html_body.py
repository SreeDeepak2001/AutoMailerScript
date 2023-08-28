def HTML_BODY(tableContent):
    htmlBody =  """
                <header>The Following Tickets need attention</header>
                <body style = "justify-content:center;">
                <br>
                <table border = 1 style = "text-align: center;">
                    <tr style = "background-color: #f0831d;">
                        <th>Ticket Number</th>
                        <th>Priority</th>
                        <th>Assignee</th>
                        <th>Status</th>
                        <th>Last Comment Date</th>
                    </tr>
                    """+tableContent+"""
                </table>
                <br><br><br>
                <img src="cid:logo_img">
                <br>
                <a href = https://www.flatironssolutions.com>https://www.flatironssolutions.com</a>
                </body>
                """

    return htmlBody

