<%@ Language=VBScript CodePage=65001 %>
<%
' -------------------------------------------------------------------------
'  SPLINE MASTER PRO - SERVER SIDE ENGINE
' -------------------------------------------------------------------------
Response.CharSet = "utf-8"

Dim conn, rs, dbPath, sql, mode, statusMsg, statusType
Dim x(5), y(5), rx(5), ry(5), tension, curveType
Dim i, t, t2, t3, h1, h2, h3, h4, xx, yy, b1, b2, b3, b4, oneOver6

' Helper: Safe Float Conversion for Globalization
Function GetFloat(val)
    If IsNull(val) Or val = "" Then GetFloat = 0 Else GetFloat = CDbl(Replace(val, ".", ","))
End Function
Function ToJS(val)
    ToJS = Replace(CStr(val), ",", ".")
End Function

' Database Connection
Set conn = Server.CreateObject("ADODB.Connection")
dbPath = Server.MapPath("splines.mdb")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath

mode = Request.Form("mode")

' Default Initialization
x(1) = 50 : y(1) = 350 : x(2) = 150 : y(2) = 50 
x(3) = 450 : y(3) = 50 : x(4) = 550 : y(4) = 350
tension = -0.5 : curveType = "cardinal"
statusMsg = "System Ready. Waiting for input..."
statusType = "info"

' Logic Flow
If mode = "calc" Or mode = "save" Then
    x(1) = GetFloat(Request.Form("x1")) : y(1) = GetFloat(Request.Form("y1"))
    x(2) = GetFloat(Request.Form("x2")) : y(2) = GetFloat(Request.Form("y2"))
    x(3) = GetFloat(Request.Form("x3")) : y(3) = GetFloat(Request.Form("y3"))
    x(4) = GetFloat(Request.Form("x4")) : y(4) = GetFloat(Request.Form("y4"))
    tension = GetFloat(Request.Form("tension"))
    curveType = Request.Form("curveType")
    
    If mode = "save" Then
        Dim sName : sName = Replace(Request.Form("saveName"), "'", "")
        If sName <> "" Then
            sql = "INSERT INTO SavedCurves (SaveName, CurveType, Tension, X1, Y1, X2, Y2, X3, Y3, X4, Y4) VALUES ("
            sql = sql & "'" & sName & "', '" & curveType & "', " & Replace(tension, ",", ".") & ", "
            sql = sql & Replace(x(1),",",".") & "," & Replace(y(1),",",".") & "," & Replace(x(2),",",".") & "," & Replace(y(2),",",".") & ","
            sql = sql & Replace(x(3),",",".") & "," & Replace(y(3),",",".") & "," & Replace(x(4),",",".") & "," & Replace(y(4),",",".") & ")"
            conn.Execute sql
            statusMsg = "Success: Configuration '" & sName & "' saved to database."
            statusType = "success"
        Else
            statusMsg = "Error: Please enter a name for the save file."
            statusType = "error"
        End If
    Else
        statusMsg = "Calculated: Scene updated based on new coordinates."
        statusType = "info"
    End If
ElseIf mode = "load" Then
    Dim lID : lID = Request.Form("loadID")
    If lID <> "" Then
        Set rs = conn.Execute("SELECT * FROM SavedCurves WHERE ID=" & lID)
        If Not rs.EOF Then
            curveType = rs("CurveType") : tension = rs("Tension")
            x(1)=rs("X1"):y(1)=rs("Y1"):x(2)=rs("X2"):y(2)=rs("Y2")
            x(3)=rs("X3"):y(3)=rs("Y3"):x(4)=rs("X4"):y(4)=rs("Y4")
            statusMsg = "Loaded: Configuration '" & rs("SaveName") & "' restored."
            statusType = "success"
        End If
        rs.Close
    End If
End If
%>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Spline Master Pro</title>
    <link rel="stylesheet" type="text/css" href="style.css">
    
    <% 
        Dim developerName, githubLink
        developerName = "Muhammed Furkan Ertal" 
        githubLink = "https://github.com/furkanertal/spline-master-pro"
    %>

    <script>
        var ctx, canvas, isDragging = false, selectedPointIndex = -1;
        var points = [ {x:0,y:0}, {x:<%=ToJS(x(1))%>,y:<%=ToJS(y(1))%>}, {x:<%=ToJS(x(2))%>,y:<%=ToJS(y(2))%>}, {x:<%=ToJS(x(3))%>,y:<%=ToJS(y(3))%>}, {x:<%=ToJS(x(4))%>,y:<%=ToJS(y(4))%>} ];

        function init() {
            canvas = document.getElementById("myCanvas");
            if(canvas.getContext){
                ctx = canvas.getContext("2d");
                canvas.addEventListener('mousedown', onMouseDown);
                canvas.addEventListener('mousemove', onMouseMove);
                canvas.addEventListener('mouseup', onMouseUp);
                canvas.addEventListener('mouseout', onMouseUp);
                renderScene();
            }
        }

        function renderScene() {
            ctx.clearRect(0,0,canvas.width,canvas.height);
            // Grid Drawing
            ctx.strokeStyle="#f3f4f6"; ctx.beginPath();
            for(var gx=0;gx<=canvas.width;gx+=50){ctx.moveTo(gx,0);ctx.lineTo(gx,canvas.height);}
            for(var gy=0;gy<=canvas.height;gy+=50){ctx.moveTo(0,gy);ctx.lineTo(canvas.width,gy);}
            ctx.stroke();

            // Render Curve (Server Logic)
            if(!isDragging){
                <% 
                Select Case curveType
                    Case "cardinal" : Call DrawCardinal()
                    Case "bezier"   : Call DrawBezier()
                    Case "bspline"  : Call DrawBSpline()
                    Case "linear"   : Call DrawLinear()
                End Select
                %>
            }

            // Render Points (Client Logic)
            for(var i=1; i<=4; i++){
                var color = (i===selectedPointIndex) ? "#f59e0b" : "#ef4444";
                var size = (i===selectedPointIndex) ? 9 : 6;
                var shadow = (i===selectedPointIndex) ? 10 : 0;
                
                ctx.shadowBlur = shadow; ctx.shadowColor = "rgba(0,0,0,0.2)";
                ctx.beginPath(); ctx.arc(points[i].x, points[i].y, size, 0, 2*Math.PI); ctx.fillStyle=color; ctx.fill();
                ctx.shadowBlur = 0;

                ctx.fillStyle="#6b7280"; ctx.font="bold 11px Inter"; ctx.fillText("P"+i, points[i].x+12, points[i].y-12);
            }
        }

        function onMouseDown(e){
            var pos = getPos(e);
            for(var i=1; i<=4; i++){
                if(Math.hypot(pos.x-points[i].x, pos.y-points[i].y) < 15){ isDragging=true; selectedPointIndex=i; break; }
            }
        }
        function onMouseMove(e){
            if(isDragging){
                var pos = getPos(e);
                points[selectedPointIndex].x = pos.x; points[selectedPointIndex].y = pos.y;
                document.getElementById("x"+selectedPointIndex).value = Math.round(pos.x);
                document.getElementById("y"+selectedPointIndex).value = Math.round(pos.y);
                renderScene();
            }
        }
        function onMouseUp(e){ if(isDragging){ isDragging=false; selectedPointIndex=-1; submitForm('calc'); } }
        function getPos(e){ var r=canvas.getBoundingClientRect(); return {x:e.clientX-r.left, y:e.clientY-r.top}; }
        function submitForm(m){ document.getElementById("modeInput").value=m; document.getElementById("mainForm").submit(); }
        
        // JS Drawing Helpers for VBScript
        function drawPoint(x,y,c,s){ ctx.beginPath(); ctx.arc(x,y,s,0,2*Math.PI); ctx.fillStyle=c; ctx.fill(); }
        function drawLine(x1,y1,x2,y2,c,w,d){ 
            ctx.beginPath(); ctx.strokeStyle=c; ctx.lineWidth=w; 
            if(d)ctx.setLineDash([5,5]); else ctx.setLineDash([]); 
            ctx.moveTo(x1,y1); ctx.lineTo(x2,y2); ctx.stroke(); ctx.setLineDash([]); 
        }
    </script>
</head>
<body onload="init()">

    <header class="app-header">
        <div class="brand">
            <h1>Spline Master</h1>
            <span>PRO</span>
        </div>
        <div class="developer-profile">
            <span>Developer: <strong><%=developerName%></strong></span>
            <a href="<%=githubLink%>" target="_blank" class="github-btn">
                <svg height="20" width="20" viewBox="0 0 16 16" fill="white"><path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"></path></svg>
                Github
            </a>
        </div>
    </header>

    <div class="container">
        
        <div class="panel-controls">
            
            <div style="padding:12px; border-radius:8px; font-size:0.85rem; font-weight:500; border:1px solid; display:flex; align-items:center; gap:8px;
                <% If statusType="success" Then %>background:#ecfdf5; color:#065f46; border-color:#a7f3d0;
                <% ElseIf statusType="error" Then %>background:#fef2f2; color:#991b1b; border-color:#fecaca;
                <% Else %>background:#eff6ff; color:#1e40af; border-color:#dbeafe;<% End If %>">
                <span>
                    <% If statusType="success" Then %>✓<% ElseIf statusType="error" Then %>✕<% Else %>ℹ<% End If %>
                </span>
                <%=statusMsg%>
            </div>

            <form id="mainForm" method="post" action="">
                <input type="hidden" id="modeInput" name="mode" value="calc">

                <div class="db-box">
                    <h3 style="color:#b45309; display:flex; align-items:center; gap:5px;">
                        <svg width="14" height="14" fill="currentColor" viewBox="0 0 16 16"><path d="M2.5 1h11a1.5 1.5 0 0 1 1.5 1.5v11a1.5 1.5 0 0 1-1.5 1.5h-11A1.5 1.5 0 0 1 1 14v-11A1.5 1.5 0 0 1 2.5 1zm0-1a2.5 2.5 0 0 0-2.5 2.5v11A2.5 2.5 0 0 0 2.5 16h11a2.5 2.5 0 0 0 2.5-2.5v-11A2.5 2.5 0 0 0 13.5 0h-11z"/><path d="M4 11.5a.5.5 0 0 1 .5-.5h7a.5.5 0 0 1 0 1h-7a.5.5 0 0 1-.5-.5zm0-2a.5.5 0 0 1 .5-.5h7a.5.5 0 0 1 0 1h-7a.5.5 0 0 1-.5-.5zm0-2a.5.5 0 0 1 .5-.5h7a.5.5 0 0 1 0 1h-7a.5.5 0 0 1-.5-.5zm0-2a.5.5 0 0 1 .5-.5h7a.5.5 0 0 1 0 1h-7a.5.5 0 0 1-.5-.5z"/></svg>
                        Database Manager
                    </h3>
                    <div class="db-row">
                        <input type="text" name="saveName" placeholder="Configuration Name..." style="background:white;">
                        <button type="button" class="btn-save" onclick="submitForm('save')">Save</button>
                    </div>
                    <div class="db-row">
                        <select name="loadID" style="background:white;">
                            <option value="">-- Load Configuration --</option>
                            <%
                            Set rs = conn.Execute("SELECT ID, SaveName FROM SavedCurves ORDER BY ID DESC")
                            Do While Not rs.EOF
                                Response.Write "<option value='" & rs("ID") & "'>" & rs("SaveName") & "</option>"
                                rs.MoveNext
                            Loop
                            rs.Close
                            %>
                        </select>
                        <button type="button" class="btn-load" onclick="submitForm('load')">Load</button>
                    </div>
                </div>

                <div style="margin-top:20px;">
                    <label>Interpolation Algorithm</label>
                    <select name="curveType" onchange="submitForm('calc')">
                        <option value="cardinal" <%If curveType="cardinal" Then Response.Write "selected"%>>Cardinal Spline (Hermite)</option>
                        <option value="bezier" <%If curveType="bezier" Then Response.Write "selected"%>>Cubic Bezier Curve</option>
                        <option value="bspline" <%If curveType="bspline" Then Response.Write "selected"%>>Uniform Cubic B-Spline</option>
                        <option value="linear" <%If curveType="linear" Then Response.Write "selected"%>>Linear Interpolation (Debug)</option>
                    </select>
                </div>

                <div style="margin-top:10px;">
                    <label>Control Point Coordinates</label>
                    <div class="coord-grid">
                        <% For i=1 To 4 %>
                        <div class="coord-item">
                            <strong>P<%=i%></strong>
                            <input type="hidden" id="x<%=i%>" name="x<%=i%>" value="<%=ToJS(x(i))%>">
                            <input type="hidden" id="y<%=i%>" name="y<%=i%>" value="<%=ToJS(y(i))%>">
                            <span><%=Int(x(i))%>, <%=Int(y(i))%></span>
                        </div>
                        <% Next %>
                    </div>
                </div>

                <div style="margin-top:10px;">
                    <label>Tension (Cardinal Only)</label>
                    <input type="text" name="tension" value="<%=ToJS(tension)%>">
                </div>
            </form>

            <div class="tutorial-section">
                <div class="tutorial-title">
                    <svg width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M8 15A7 7 0 1 1 8 1a7 7 0 0 1 0 14zm0 1A8 8 0 1 0 8 0a8 8 0 0 0 0 16z"/><path d="M5.255 5.786a.237.237 0 0 0 .241.247h.825c.138 0 .248-.113.266-.25.09-.656.54-1.134 1.342-1.134.686 0 1.314.343 1.314 1.168 0 .635-.374.927-.965 1.371-.673.489-1.206 1.06-1.168 1.987l.003.217a.25.25 0 0 0 .25.246h.811a.25.25 0 0 0 .25-.25v-.105c0-.718.273-.927 1.01-1.486.609-.463 1.244-.977 1.244-2.056 0-1.511-1.276-2.241-2.673-2.241-1.267 0-2.655.59-2.75 2.286zm1.557 5.763c0 .533.425.927 1.01.927.609 0 1.028-.394 1.028-.927 0-.552-.42-.94-1.029-.94-.584 0-1.009.388-1.009.94z"/></svg>
                    How to use
                </div>
                <div class="tutorial-text">
                    This engine interpolates 4 control points using various mathematical models.
                    <ul>
                        <li><strong>Cardinal:</strong> Interpolation. Passes through all points. Use 'Tension' to adjust tightness (Catmull-Rom is Tension 0.5).</li>
                        <li><strong>Bezier:</strong> Approximation. P2 & P3 are control handles, curve does not pass through them.</li>
                        <li><strong>B-Spline:</strong> Approximation. Max smoothness (C2 continuity).</li>
                        <li><strong>Linear:</strong> Basic polygonal chain. Useful for debugging control polygon.</li>
                    </ul>
                    <em>Tip: Drag and drop the points to update the math engine in real-time.</em>
                </div>
            </div>

        </div>

        <div class="panel-canvas">
            <canvas id="myCanvas" width="800" height="500"></canvas>
        </div>
    </div>
</body>
</html>

<%
' =========================================================
' VBSCRIPT MATHEMATICAL LIBRARY
' =========================================================

' 1. CARDINAL SPLINE
Sub DrawCardinal()
    rx(1)=0.5*(1-tension)*(x(2)-x(1)):ry(1)=0.5*(1-tension)*(y(2)-y(1))
    rx(2)=0.5*(1-tension)*(x(3)-x(1)):ry(2)=0.5*(1-tension)*(y(3)-y(1))
    rx(3)=0.5*(1-tension)*(x(4)-x(2)):ry(3)=0.5*(1-tension)*(y(4)-y(2))
    rx(4)=0.5*(1-tension)*(x(4)-x(3)):ry(4)=0.5*(1-tension)*(y(4)-y(3))
    For i=1 To 3
        For t=0 To 1 Step 0.02
            t2=t*t:t3=t*t*t
            h1=(2*t3)-(3*t2)+1 : h2=(-2*t3)+(3*t2) : h3=t3-(2*t2)+t : h4=t3-t2
            xx=x(i)*h1+x(i+1)*h2+rx(i)*h3+rx(i+1)*h4
            yy=y(i)*h1+y(i+1)*h2+ry(i)*h3+ry(i+1)*h4
            Response.Write "drawPoint("&ToJS(xx)&","&ToJS(yy)&",'#4f46e5',1.5);"&vbCrLf
        Next
    Next
End Sub

' 2. CUBIC BEZIER
Sub DrawBezier()
    Response.Write "drawLine("&ToJS(x(1))&","&ToJS(y(1))&","&ToJS(x(2))&","&ToJS(y(2))&",'#9ca3af',1,true);"
    Response.Write "drawLine("&ToJS(x(3))&","&ToJS(y(3))&","&ToJS(x(4))&","&ToJS(y(4))&",'#9ca3af',1,true);"
    For t=0 To 1 Step 0.01
        t2=t*t:t3=t*t*t
        b1=(1-t)*(1-t)*(1-t) : b2=3*(1-t)*(1-t)*t : b3=3*(1-t)*t*t : b4=t3
        xx=x(1)*b1+x(2)*b2+x(3)*b3+x(4)*b4
        yy=y(1)*b1+y(2)*b2+y(3)*b3+y(4)*b4
        Response.Write "drawPoint("&ToJS(xx)&","&ToJS(yy)&",'#10b981',1.5);"&vbCrLf
    Next
End Sub

' 3. UNIFORM B-SPLINE
Sub DrawBSpline()
    Response.Write "drawLine("&ToJS(x(1))&","&ToJS(y(1))&","&ToJS(x(2))&","&ToJS(y(2))&",'#d1d5db',1,true);"
    Response.Write "drawLine("&ToJS(x(2))&","&ToJS(y(2))&","&ToJS(x(3))&","&ToJS(y(3))&",'#d1d5db',1,true);"
    Response.Write "drawLine("&ToJS(x(3))&","&ToJS(y(3))&","&ToJS(x(4))&","&ToJS(y(4))&",'#d1d5db',1,true);"
    oneOver6=1.0/6.0
    For t=0 To 1 Step 0.01
        t2=t*t:t3=t*t*t
        b1=oneOver6*((1-t)*(1-t)*(1-t)) : b2=oneOver6*(3*t3-6*t2+4)
        b3=oneOver6*(-3*t3+3*t2+3*t+1) : b4=oneOver6*t3
        xx=x(1)*b1+x(2)*b2+x(3)*b3+x(4)*b4
        yy=y(1)*b1+y(2)*b2+y(3)*b3+y(4)*b4
        Response.Write "drawPoint("&ToJS(xx)&","&ToJS(yy)&",'#8b5cf6',1.5);"&vbCrLf
    Next
End Sub

' 4. LINEAR INTERPOLATION (NEW FEATURE)
Sub DrawLinear()
    Response.Write "drawLine("&ToJS(x(1))&","&ToJS(y(1))&","&ToJS(x(2))&","&ToJS(y(2))&",'#3b82f6',2,false);"
    Response.Write "drawLine("&ToJS(x(2))&","&ToJS(y(2))&","&ToJS(x(3))&","&ToJS(y(3))&",'#3b82f6',2,false);"
    Response.Write "drawLine("&ToJS(x(3))&","&ToJS(y(3))&","&ToJS(x(4))&","&ToJS(y(4))&",'#3b82f6',2,false);"
End Sub

conn.Close : Set conn = Nothing

%>
