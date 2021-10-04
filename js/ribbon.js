
//这个函数在整个wps加载项中是第一个执行的
function OnAddinLoad(ribbonUI){
    if (typeof (wps.ribbonUI) != "object"){
		wps.ribbonUI = ribbonUI
    }
    
    if (typeof (wps.Enum) != "object") { // 如果没有内置枚举值
        wps.Enum = WPS_Enum
    }

    wps.PluginStorage.setItem("EnableFlag", false) //往PluginStorage中设置一个标记，用于控制两个按钮的置灰
    wps.PluginStorage.setItem("ApiEventFlag", false) //往PluginStorage中设置一个标记，用于控制ApiEvent的按钮label
    
    return true
}

var WebNotifycount = 0;
function OnAction(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            {
                const doc = wps.WpsApplication().ActiveDocument
                if (!doc) {
                    alert("当前没有打开任何文档")
                    return
                }
                alert("hello world")
            }
            break;
        case "btnIsEnbable":
            {
                let bFlag = wps.PluginStorage.getItem("EnableFlag")
                wps.PluginStorage.setItem("EnableFlag", !bFlag)
                
                //通知wps刷新以下几个按饰的状态
                wps.ribbonUI.InvalidateControl("btnIsEnbable")
                wps.ribbonUI.InvalidateControl("btnShowDialog") 
                wps.ribbonUI.InvalidateControl("btnShowTaskPane") 
                //wps.ribbonUI.Invalidate(); 这行代码打开则是刷新所有的按钮状态
                break
            }
        case "btnShowDialog":
            wps.ShowDialog(GetUrlPath() + "/ui/dialog.html", "新建人物角色", 400 * window.devicePixelRatio, 400 * window.devicePixelRatio, false)
            break
        case "btnShowTaskPane":
            {
                let tsId = wps.PluginStorage.getItem("taskpane_id")
                if (!tsId) {
                    let tskpane = wps.CreateTaskPane(GetUrlPath() + "/ui/taskpane.html")
                    let id = tskpane.ID
                    wps.PluginStorage.setItem("taskpane_id", id)
                    tskpane.Visible = true
                } else {
                    let tskpane = wps.GetTaskPane(tsId)
                    tskpane.Visible = !tskpane.Visible
                }
            }
            break
        case "btnApiEvent":
            {
                let bFlag = wps.PluginStorage.getItem("ApiEventFlag")
                let bRegister = bFlag ? false : true
                wps.PluginStorage.setItem("ApiEventFlag", bRegister)
                if (bRegister){
                    wps.ApiEvent.AddApiEventListener('DocumentNew', OnNewDocumentApiEvent)
                }
                else{
                    wps.ApiEvent.RemoveApiEventListener('DocumentNew', OnNewDocumentApiEvent)
                }

                wps.ribbonUI.InvalidateControl("btnApiEvent") 
            }
            break
        case "btnWebNotify":
            {
                let currentTime = new Date()
                let timeStr = currentTime.getHours() + ':' + currentTime.getMinutes() + ":" + currentTime.getSeconds()
                wps.OAAssist.WebNotify("这行内容由wps加载项主动送达给业务系统，可以任意自定义, 比如时间值:" + timeStr + "，次数：" + (++WebNotifycount), true)
            }
            break
        case "test1":
            {
                if (Application.Selection.Font.Bold == 0){
                    Application.Selection.Font.Bold = -1
                } else {
                    Application.Selection.Font.Bold = 0
                }
                      alert(Application.Selection.Text)   

                      
                         
            }
            break
        case "test2":
            {
                // 扩展到整段
                Application.Selection.Expand(4)   
                //Application.Selection.Range.HighlightColorIndex = 4;
                //Application.Selection.Range.HighlightColorIndex=3
                //黑色默认全0
                console.log("get hex:%s\n",(0x1000000+ Application.Selection.Font.Color).toString(16).substring(1));
                // alert(Application.Selection.Font.Color) 
                // if(Application.Selection.Font.Color = 0xffffff){
                //     alert("default")    
                // }
                
                Application.Selection.Font.Color = 0xff0000
                //Selection.MoveLeft(1,1,0);
                //Selection.TypeText("【test Masg】");
                // Application.Selection.HomeKey(0,1)
                // Application.Selection.TypeText("【test Masg】")
                Application.Selection.InsertBefore("\"Hamlet\"")
                
            }
            break
        default:
            break
    }
    return true
}

function GetImage(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            return "images/1.svg"
        case "btnShowDialog":
            return "images/2.svg"
        case "btnShowTaskPane":
            return "images/3.svg"
        default:
            ;
    }
    return "images/newFromTemp.svg"
}

function OnGetEnabled(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            return true
            break
        case "btnShowDialog":
            {
                let bFlag = wps.PluginStorage.getItem("EnableFlag")
                return bFlag
                break
            }
        case "btnShowTaskPane":
            {
                let bFlag = wps.PluginStorage.getItem("EnableFlag")
                return bFlag
                break
            }
        case "New_Chara":{
            let bFlag = wps.PluginStorage.getItem("EnableFlag")
            return bFlag
            break
        }
        case "btnShowTaskPane":{
            let bFlag = wps.PluginStorage.getItem("EnableFlag")
            return bFlag
            break
        }

        default:
            break
    }
    return true
}

function OnGetVisible(control){
    return true
}

function OnGetLabel(control){
    const eleId = control.Id
    switch (eleId) {
        case "btnIsEnbable":
        {
            let bFlag = wps.PluginStorage.getItem("EnableFlag")
            return bFlag ?  "按钮Disable" : "按钮Enable"
            break
        }
        case "btnApiEvent":
        {
            let bFlag = wps.PluginStorage.getItem("ApiEventFlag")
            return bFlag ? "清除新建文件事件" : "注册新建文件事件"
            break
        }    
    }
    return ""
}

function OnNewDocumentApiEvent(doc){
    alert("新建文件事件响应，取文件名: " + doc.Name)
}


function WAMPO_init(){
    console.log(Application.ActiveDocument.Tables.Item(1).Columns.Count)
    if (Application.ActiveDocument.Tables.Item(1)){
        // var chara_table = Application.ActiveDocument.Tables.Item(1)
        // console.log(chara_table.Cell(1,1).Width=20)
        if(Application.ActiveDocument.Tables.Item(1).Columns.Count == 5){

            var chara_table = Application.ActiveDocument.Tables.Item(1)
            chara_table.Cell(1,1).Range.Select()
            var header1 = Application.Selection.Text.replace( /^\s+|\s+$/g, "" );
            chara_table.Cell(1,2).Range.Select()
            var header2 = Application.Selection.Text.replace( /^\s+|\s+$/g, "" );
            chara_table.Cell(1,3).Range.Select()
            var header3 = Application.Selection.Text.replace( /^\s+|\s+$/g, "" );
            chara_table.Cell(1,4).Range.Select()
            var header4 = Application.Selection.Text.replace( /^\s+|\s+$/g, "" );
            chara_table.Cell(1,5).Range.Select()
            var header5 = Application.Selection.Text.replace( /^\s+|\s+$/g, "" );

            if (header1=="角色名称" && header2=="CV名称" && header3=="角色描述" && header4=="角色性别" && header5=="着色预览"){
                Application.Selection.EndKey(Application.Enum.wdLine,Application.Enum.wdMove)
                return true    
            } else {
                var con = confirm("人物属性表存在问题，无法正确解析，是否重新创建？");
                if(con == true){
                    Construct_Main_Form();
                    Application.Selection.EndKey(Application.Enum.wdLine,Application.Enum.wdMove)
                    return true
                } 
            }
        } else {
                var con = confirm("人物属性表存在问题，无法正确解析，是否重新创建？");
                if(con == true){
                    Construct_Main_Form();
                    Application.Selection.EndKey(Application.Enum.wdLine,Application.Enum.wdMove)
                    return true
                } 
            }
    } else { 

        var con = confirm("未检测到人物属性表，是否立即创建？");
        if(con == true){
            Construct_Main_Form();
            Application.Selection.EndKey(Application.Enum.wdLine,Application.Enum.wdMove)
            return true
        }
        return false

    }
}

function Construct_Main_Form(){

    Application.ActiveDocument.Range(0, 0).InsertBreak()
    Application.ActiveDocument.Range(0, 0).Select()
    Application.Selection.InsertBefore("* 此表格为画本工具自动生成，包含关键配置，任意修改可能导致程序混乱")
    Application.Selection.Expand(4)
    Application.Selection.Font.Size = 8
    Application.Selection.Font.Bold = -1
    Application.Selection.ParagraphFormat.Alignment = Application.Enum.wdAlignParagraphLeft;
    Application.Selection.Font.Color = 0x0

    let tblNew = Application.ActiveDocument.Tables.Add(Application.ActiveDocument.Range(0, 0), 1, 5)
    tblNew.Cell(1, 1).Range.InsertAfter("角色名称")
    tblNew.Cell(1, 1).Range.Select()
    Application.Selection.Font.Bold = 0
    Application.Selection.Font.Color = 0x666666 
    Application.Selection.Font.Size = 10
    Application.Selection.ParagraphFormat.Alignment = Application.Enum.wdAlignParagraphCenter;

    tblNew.Cell(1, 2).Range.InsertAfter("CV名称")
    tblNew.Cell(1, 2).Range.Select()
    Application.Selection.Font.Color = 0x666666 
    Application.Selection.Font.Size = 10
    Application.Selection.ParagraphFormat.Alignment = Application.Enum.wdAlignParagraphCenter;

    tblNew.Cell(1, 3).Range.InsertAfter("角色描述")
    tblNew.Cell(1, 3).Range.Select()
    Application.Selection.Font.Color = 0x666666 
    Application.Selection.Font.Size = 10
    Application.Selection.ParagraphFormat.Alignment = Application.Enum.wdAlignParagraphCenter;

    tblNew.Cell(1, 4).Range.InsertAfter("角色性别")
    tblNew.Cell(1, 4).Range.Select()
    Application.Selection.Font.Color = 0x666666 
    Application.Selection.Font.Size = 10
    Application.Selection.ParagraphFormat.Alignment = Application.Enum.wdAlignParagraphCenter;

    tblNew.Cell(1, 5).Range.InsertAfter("着色预览")
    tblNew.Cell(1, 5).Range.Select()
    Application.Selection.Font.Color = 0x666666 
    Application.Selection.Font.Size = 10
    Application.Selection.ParagraphFormat.Alignment = Application.Enum.wdAlignParagraphCenter;

    // tblNew.Columns.AutoFit()
    let myTable = Application.ActiveDocument.Tables.Item(1)
    myTable.Borders.InsideLineStyle = Application.Enum.wdLineStyleDashLargeGap
    myTable.Borders.OutsideLineStyle = Application.Enum.wdLineStyleSingle

    myTable.Cell(1,1).Width=60
    myTable.Cell(1,2).Width=60
    myTable.Cell(1,3).Width=200
    myTable.Cell(1,4).Width=60
    myTable.Cell(1,5).Width=60

}

function WAMPO_MAIN_ACTION(control){
    const eleId = control.Id
    switch (eleId) {
        case "Startup":
        {
            const doc = wps.WpsApplication().ActiveDocument
            if (!doc) {
                alert("当前没有打开任何文档")
                return
            } else {
                
                let tsId = wps.PluginStorage.getItem("taskpane_id")
                if (!tsId) {
                    

                    let tskpane = wps.CreateTaskPane(GetUrlPath() + "/ui/sidebar.html")
                    let id = tskpane.ID
                    wps.PluginStorage.setItem("taskpane_id", id)
                } 

                let tskpane = wps.GetTaskPane(tsId)
                if(!tskpane.Visible){
                    WAMPO_init()
                    wps.PluginStorage.setItem("EnableFlag", true)                  

                    wps.ribbonUI.InvalidateControl("New_Chara")


                    tskpane.Visible = true
                    tskpane.Navigate(GetUrlPath() + "/ui/sidebar.html")
                } else {
                    wps.PluginStorage.setItem("EnableFlag", false)
                    wps.ribbonUI.InvalidateControl("New_Chara")
                    tskpane.Visible = false
                }


            }
            break
        }
        case "New_Chara":
        {
            wps.ShowDialog(GetUrlPath() + "/ui/dialog.html", "新建人物属性", 400 * window.devicePixelRatio, 400 * window.devicePixelRatio, false)
            break

        }
         

        default:
            break

    }
    return true


}