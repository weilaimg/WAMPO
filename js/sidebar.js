var myCS_Name = new Array();
var myCV_Name = new Array();
var myF_Color = new Array();
var myB_Color = new Array();

window.onload = function () {
    // var xmlReq = WpsInvoke.CreateXHR();
    // var url = location.origin + "/.debugTemp/NotifyDemoUrl"
    // xmlReq.open("GET", url);
    // xmlReq.onload = function (res) {
    //     var node = document.getElementById("DemoSpan");
    //     node.innerText = res.target.responseText;
    // };
    // xmlReq.send();


    var row_num = Application.ActiveDocument.Tables.Item(1).Rows.Count

    for (var i = 2; i<=row_num; i++){
        

        var table_div = document.getElementById("Chara_Table")

        var tr_div = document.createElement("div");
        tr_div.setAttribute("class","table-tr")
        tr_div.setAttribute("onclick","onbuttonclick('row\,"+i+"')")
        
        Application.ActiveDocument.Tables.Item(1).Cell(i,1).Range.Select()
        var td_div = document.createElement("div");
        td_div.setAttribute("class","table-td")
        var CS_Name =Application.Selection.Text.replace( /^\s+|\s+$/g, "" )
        td_div.innerHTML=CS_Name
        myCS_Name[i] = CS_Name
        tr_div.appendChild(td_div);

        Application.ActiveDocument.Tables.Item(1).Cell(i,2).Range.Select()
        var td_div = document.createElement("div");
        td_div.setAttribute("class","table-td")
        var CV_Name = Application.Selection.Text.replace( /^\s+|\s+$/g, "" )
        td_div.innerHTML= CV_Name
        myCV_Name[i] = CV_Name
        tr_div.appendChild(td_div);

        Application.ActiveDocument.Tables.Item(1).Cell(i,4).Range.Select()
        var td_div = document.createElement("div");
        td_div.setAttribute("class","table-td")
        td_div.innerHTML=Application.Selection.Text.replace( /^\s+|\s+$/g, "" )
        tr_div.appendChild(td_div);


        Application.ActiveDocument.Tables.Item(1).Cell(i,5).Range.Select()
        var Font_B = Application.Selection.Font.Color>>16
        var Font_G = Application.Selection.Font.Color<<8>>16&0xff
        var Font_R = Application.Selection.Font.Color<<16>>16&0xff
        myF_Color[i] = Application.Selection.Font.Color

        Application.ActiveDocument.Tables.Item(1).Cell(i,5).Select()
        Application.Selection.HomeKey(Application.Enum.wdLine,Application.Enum.wdMove)
        Application.Selection.Expand(2)
        
        var Highlight_Color = Get_HighlightColor(Application.Selection.Range.HighlightColorIndex)
        var Highlight_R = Highlight_Color>>16
        var Highlight_G = Highlight_Color<<8>>16&0xff
        var Highlight_B = Highlight_Color<<16>>16&0xff
        myB_Color[i] = Application.Selection.Range.HighlightColorIndex

        var td_div = document.createElement("div");
        td_div.setAttribute("class","table-td")

        var color =  "color: rgb("+Font_R+","+Font_G+","+Font_B+");"+"background-color: rgb("+Highlight_R+","+Highlight_G+","+Highlight_B+");"
        td_div.setAttribute("style",color);

        td_div.innerHTML=CS_Name
        tr_div.appendChild(td_div);

        table_div.appendChild(tr_div);  

    }

    Application.Selection.EndKey(Application.Enum.wdLine,Application.Enum.wdMove)

}




function onbuttonclick(idStr)
{
    // if (typeof (wps.Enum) != "object") { // 如果没有内置枚举值
    //     wps.Enum = WPS_Enum
    // }


    sw_str = idStr.split(",")[0]
    switch(sw_str)
    {
        
        case 'refresh':{
            location.reload();
            break
        }
        case 'row':{
            row_num = Number(idStr.split(",")[1])
            do_paint(row_num)
            
            break
        }
    }

}

function do_paint(row_num){
    // let SelLength = 
    if (Application.Selection.End - Application.Selection.Start){
        Application.Selection.Range.HighlightColorIndex = myB_Color[row_num]
        Application.Selection.Font.Color = myF_Color[row_num]

        Application.Selection.MoveLeft(1,1,0);
        Input_Text = "【"+myCS_Name[row_num]+"】"
        Application.Selection.TypeText(Input_Text);
        Application.Selection.SetRange(Application.Selection.End-Input_Text.length,Application.Selection.End);
        Application.Selection.Range.HighlightColorIndex = myB_Color[row_num]
        Application.Selection.Font.Color = myF_Color[row_num]
        Application.Selection.MoveLeft(1,1,0);


    } else {

        Application.Selection.Expand(4)
        Application.Selection.Range.HighlightColorIndex = myB_Color[row_num]
        Application.Selection.Font.Color = myF_Color[row_num]

        Application.Selection.MoveLeft(1,1,0);
        Application.Selection.TypeText("【"+myCS_Name[row_num]+"】");


    }
    
    // Application.ActiveDocument.Tables.Item(1).Cell(row_num,5).Range.Select()
    // var Font_Color = Application.Selection.Font.Color

    // Application.ActiveDocument.Tables.Item(1).Cell(row_num,5).Select()
    // Application.Selection.HomeKey(Application.Enum.wdLine,Application.Enum.wdMove)
    // Application.Selection.Expand(2)
    // var Ht_Color_Index = Application.Selection.Range.HighlightColorIndex

    

}