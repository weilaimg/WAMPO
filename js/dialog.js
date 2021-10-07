function onbuttonclick(idStr) {
    switch (idStr) {
        case "getDocName":
            {
                let doc = wps.WpsApplication().ActiveDocument
                let textValue = ""
                if (!doc) {
                    textValue = textValue + "当前没有打开任何文档"
                    return
                }
                textValue = textValue + doc.Name
                document.getElementById("text_p").innerHTML = textValue
                break
            }
        case "createTaskPane":
            {
                let tsId = wps.PluginStorage.getItem("taskpane_id")
                if (!tsId) {
                    let tskpane = wps.CreateTaskPane(GetUrlPath() + "/taskpane.html")
                    let id = tskpane.ID
                    wps.PluginStorage.setItem("taskpane_id", id)
                    tskpane.Visible = true
                } else {
                    let tskpane = wps.GetTaskPane(tsId)
                    tskpane.Visible = true
                }
                break
            }
        case "newDoc":
            {
                wps.WpsApplication().Documents.Add()
                break
            }
        case "addString":
            {
                let doc = wps.WpsApplication().ActiveDocument
                if (doc) {
                    doc.Range(0, 0).Text = "Hello, wps加载项!"
                    //好像是wps的bug, 这两句话触发wps重绘
                    let rgSel = wps.WpsApplication().Selection.Range
                    if (rgSel)
                        rgSel.Select()
                }
                break;
            }
        case "closeDoc":
            {
                if (wps.WpsApplication().Documents.Count < 2) {
                    alert("当前只有一个文档，别关了。")
                    break
                }

                let doc = wps.WpsApplication().ActiveDocument
                if (doc)
                    doc.Close()
                break
            }
    }

}

window.onload = function() {
    // var xmlReq = WpsInvoke.CreateXHR();
    // var url = location.origin + "/.debugTemp/NotifyDemoUrl"
    // xmlReq.open("GET", url);
    // xmlReq.onload = function(res) {
    //     var node = document.getElementById("DemoSpan");
    //     node.innerText = res.target.responseText;
    // };
    // xmlReq.send();
}

colorPicker = function(idStr) {
    this.colorPool = ["#000000", "#993300", "#333300", "#003300", "#003366", "#000080", "#333399", "#333333", "#800000", "#FF6600", "#808000", "#008000", "#008080", "#0000FF", "#666699", "#808080", "#FF0000", "#FF9900", "#99CC00", "#339966", "#33CCCC", "#3366FF", "#800080", "#999999", "#FF00FF", "#FFCC00", "#FFFF00", "#00FF00", "#00FFFF", "#00CCFF", "#993366", "#CCCCCC", "#FF99CC", "#FFCC99", "#FFFF99", "#CCFFCC", "#CCFFFF", "#99CCFF", "#CC99FF", "#FFFFFF"];
    this.initialize(idStr);
}

colorPicker.prototype = {

    initialize: function(idStr) {
        var count = 0;
        var html = '';
        var self = this;
        html += '<table cellspacing="5" cellpadding="0" border="2" bordercolor="#000000" style="cursor:pointer;background:#ECE9D8" mce_style="cursor:pointer;background:#ECE9D8" >';

        for (i = 0; i < 5; i++) {
            html += "<tr>";
            for (j = 0; j < 8; j++) {
                html += '<td align="center" width="20" height="20" style="background:' + this.colorPool[count] + ';border-width: 2px; border-color:rgb(255,255,255)' + '" mce_style="background:' + this.colorPool[count] + '" unselectable="on"> </td>';
                count++;
            }
            html += "</tr>";
        }
        html += '</table>';

        this.trigger = document.getElementById(idStr);
        this.div = document.createElement('div');
        this.div.innerHTML = html;
        var tds = this.div.getElementsByTagName('td');
        for (var i = 0, l = tds.length; i < l; i++) {
            tds[i].onclick = function() {
                self.setColor(this.style.backgroundColor, idStr);
            }
        }

        this.div.id = 'myColorPicker';
        this.trigger.parentNode.appendChild(this.div);
        this.div.style.position = 'absolute';

        this.div.style.top = (this.trigger.clientHeight + this.trigger.offsetTop) + 'px';
        // this.hide();
        this.trigger.onclick = function() {
            if (self.div.style.display == 'none') {
                self.show();
                return false;
            } else {
                self.hide();
                return false;
            }
        }
    },

    setColor: function(c, idStr) {
        this.hide();
        // document.getElementById(idStr).style.backgroundColor = c //proEditor.setColor(c); //自己定义函数决定setColor的功能
        document.getElementById(idStr).style.backgroundColor = c

        //var rgb2Hex = colorRGB2Hex(c);
        //alert(rgb2Hex);
    },

    hide: function() {
        this.div.style.display = 'none'
    },

    show: function() {
        this.div.style.display = 'block'
    }

}




bgcolorPicker = function(idStr) {
    this.colorPool = new Array();
    for (var i = 1; i <= 16; i++){
        Highlight_Color=Get_HighlightColor(i)
        Font_R = Highlight_Color>>16
        Font_G = Highlight_Color<<8>>16&0xff
        Font_B = Highlight_Color<<16>>16&0xff
        this.colorPool[i-1]="rgb("+Font_R+","+Font_G+","+Font_B+")"

    }
    // this.colorPool = ["#000000", "#993300", "#333300", "#003300", "#003366", "#000080", "#333399", "#333333", "#800000", "#FF6600", "#808000", "#008000", "#008080", "#0000FF", "#666699", "#808080", "#FF0000", "#FF9900", "#99CC00", "#339966", "#33CCCC", "#3366FF", "#800080", "#999999", "#FF00FF", "#FFCC00", "#FFFF00", "#00FF00", "#00FFFF", "#00CCFF", "#993366", "#CCCCCC", "#FF99CC", "#FFCC99", "#FFFF99", "#CCFFCC", "#CCFFFF", "#99CCFF", "#CC99FF", "#FFFFFF"];
    this.initialize(idStr);
}
bgcolorPicker.prototype = {

    initialize: function(idStr) {
        var count = 1;
        var html = '';
        var self = this;
        html += '<table cellspacing="5" cellpadding="0" border="2" bordercolor="#000000" style="cursor:pointer;background:#ECE9D8" mce_style="cursor:pointer;background:#ECE9D8" >';

        for (i = 0; i < 3; i++) {
            html += "<tr>";
            for (j = 0; j < 5; j++) {
                html += '<td align="center" width="20" height="20" id="'+(count+1)+'" style="background:' + this.colorPool[count] + ';border-width: 2px; border-color:rgb(255,255,255)' + '" mce_style="background:' + this.colorPool[count] + '" unselectable="on"> </td>';
                count++;
            }
            html += "</tr>";
        }
        html += '</table>';

        this.trigger = document.getElementById(idStr);
        this.div = document.createElement('div');
        this.div.innerHTML = html;
        var tds = this.div.getElementsByTagName('td');
        for (var i = 0, l = tds.length; i < l; i++) {
            tds[i].onclick = function() {
                self.setColor(this.style.backgroundColor,this.id, idStr);
            }
        }

        this.div.id = 'myColorPicker';
        this.trigger.parentNode.appendChild(this.div);
        this.div.style.position = 'absolute';

        this.div.style.top = (this.trigger.clientHeight + this.trigger.offsetTop) + 'px';
        this.div.style.left = '170px'
        // this.hide();
        this.trigger.onclick = function() {
            if (self.div.style.display == 'none') {
                self.show();
                return false;
            } else {
                self.hide();
                return false;
            }
        }
    },

    setColor: function(c,i, idStr) {
        this.hide();
        // document.getElementById(idStr).style.backgroundColor = c //proEditor.setColor(c); //自己定义函数决定setColor的功能
        // document.getElementById(idStr).name = i
        document.getElementById(idStr).setAttribute("BGIndex",i) 
        document.getElementById(idStr).style.backgroundColor = c

        //var rgb2Hex = colorRGB2Hex(c);
        //alert(rgb2Hex);
    },

    hide: function() {
        this.div.style.display = 'none'
    },

    show: function() {
        this.div.style.display = 'block'
    }

}






function initColorPicker(str) {

    if (str == "font_color") {
        picker = new colorPicker(str);
    } else if (str == "bg_color") {
        picker = new bgcolorPicker(str);
    }

}

function colorRGB2Hex(color) {
    var rgb = color.split(',');
    var r = parseInt(rgb[0].split('(')[1]);
    var g = parseInt(rgb[1]);
    var b = parseInt(rgb[2].split(')')[0]);
    // var hex = "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1); 
    // var hex = ((r << 16) + (g << 8) + b)
    var hex = ((b << 16) + (g << 8) + r)
    return hex;

}

function btn_cancel() {

    window.close()
}


function create_chara() {
    var CS_Name = document.getElementById("CS_Name").value;
    var CV_Name = document.getElementById("CV_Name").value;
    var CV_Name_word = document.getElementById("CV_Name_word").value;
    var radio = document.getElementsByName("sex");
    var sex_value
    for (i = 0; i < radio.length; i++) {
        if (radio[i].checked) {
            sex_value = radio[i].value
        }
    }


    var font_color = document.getElementById("font_color").style.backgroundColor
    var bg_color = document.getElementById("bg_color").getAttribute("BGIndex") 
    console.log(bg_color)


    // window.close();
    // alert(Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Index)
    // alert(Application.ActiveDocument.Tables.Item(1).Columns.Count)

    if (Application.ActiveDocument.Tables.Item(1).Columns.Count == 5) {

        let tblNew = Application.ActiveDocument.Tables.Item(1)
        let rowNew = tblNew.Rows.Add(tblNew.Rows.Item(tblNew.Rows.Count + 1))
        let celTable = rowNew.Cells.Item(1)
        celTable.Range.InsertAfter(CS_Name)
        tblNew.Cell(tblNew.Rows.Count + 1, 1).Range.Select()
        Application.Selection.Font.Bold = 0
        Application.Selection.Font.Color = 0x666666
        Application.Selection.Font.Size = 10
        Application.Selection.ParagraphFormat.Alignment = Application.Enum.wdAlignParagraphCenter;

        celTable = rowNew.Cells.Item(2)
        celTable.Range.InsertAfter(CV_Name)
        tblNew.Cell(tblNew.Rows.Count + 1, 2).Range.Select()
        Application.Selection.Font.Bold = 0
        Application.Selection.Font.Color = 0x666666
        Application.Selection.Font.Size = 10
        Application.Selection.ParagraphFormat.Alignment = Application.Enum.wdAlignParagraphCenter;

        celTable = rowNew.Cells.Item(3)
        celTable.Range.InsertAfter(CV_Name_word)
        tblNew.Cell(tblNew.Rows.Count + 1, 3).Range.Select()
        Application.Selection.Font.Bold = 0
        Application.Selection.Font.Color = 0x666666
        Application.Selection.Font.Size = 10
        Application.Selection.ParagraphFormat.Alignment = Application.Enum.wdAlignParagraphCenter;

        celTable = rowNew.Cells.Item(4)
        celTable.Range.InsertAfter(sex_value)
        tblNew.Cell(tblNew.Rows.Count + 1, 4).Range.Select()
        Application.Selection.Font.Bold = 0
        Application.Selection.Font.Color = 0x666666
        Application.Selection.Font.Size = 10
        Application.Selection.ParagraphFormat.Alignment = Application.Enum.wdAlignParagraphCenter;

        celTable = rowNew.Cells.Item(5)
        celTable.Range.InsertAfter(CS_Name)
        // celTable = rowNew.Cells.Item(6)
        // celTable.Range.InsertAfter(CV_Name)
        // console.log("%s", (colorRGB2Hex(font_color) ))
        celTable.Range.Select()
        Application.Selection.Font.Color = colorRGB2Hex(font_color)
        
        Application.Selection.Range.HighlightColorIndex = bg_color

        tblNew.Cell(tblNew.Rows.Count + 1, 5).Range.Select()
        Application.Selection.Font.Bold = 0
        Application.Selection.Font.Size = 10
        // Application.Selection.Range.HighlightColorIndex = bg_color
        Application.Selection.ParagraphFormat.Alignment = Application.Enum.wdAlignParagraphCenter;


        let tsId = wps.PluginStorage.getItem("taskpane_id")
        let tskpane = wps.GetTaskPane(tsId)
        if (tskpane) {
            tskpane.Visible = true
            tskpane.Navigate(GetUrlPath() + "/sidebar.html")
        }

        Application.Selection.EndKey(Application.Enum.wdLine, Application.Enum.wdMove)
        // window.close()
    }



}


function testmouse(){
    alert('123')
}