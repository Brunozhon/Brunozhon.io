<!DOCTYPE html>
<html>
  <head>
    <title>Asp example</title>
  </head>
  <body>
    <p>
    There
    <%
    If users=1 Then
      response.write(" is: ")
    Else
      response.write(" are: ")
    End If
    %>
    </p>
    <h1>
    <%
    Dim users=Application("users")
    response.write(users)
    %>
    </h1>
    <p>
    active
    <%
    If users=1 Then
      response.write(" connection")
    Else
      response.write(" connections")
    End If
    %>
    </p>
    <script languge="vbscript" runat="server">
      Sub Application On_Start
        application("vartime")=""
        application("users")=1
      End Sub
    </script>
  </body>
</html>
