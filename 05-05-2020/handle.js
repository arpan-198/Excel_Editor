
let fun=()=>{
    localStorage.clear();
    document.getElementById('show').innerHTML="";
    let file1=document.getElementById('file-upload');
    if(file1.files[0]==null)
    {
        alert("No File Selected");
        return false;
    }
    let reader=new FileReader();
    if(reader.readAsBinaryString)
    {
        let rd=reader.readAsBinaryString(file1.files[0]);
        reader.onload=function(e){
            readExcel(e.target.result);
        };
    }
    else{
        reader.readAsArrayBuffer(file1.files[0]);
        reader.onload=function(e){
            let data="";
            let bytes=new Uint8Array(e.target.result);
            for(let i=0;i<bytes.byteLength;i++)
            data+=String.fromCharCode(bytes[i]);
        }
        readExcel(data);
    }
}

let readExcel=(bidata)=>{
    let workbook=XLSX.read(bidata,{
        type:'binary'
    });

    let ws=workbook.SheetNames;
    let wsNo=ws.length;
    localStorage.setItem("sheetNames",JSON.stringify(ws));

    document.getElementById('list').innerText="";

    let slt=document.createElement('select');
    slt.id="select1";
    var opt=document.createElement('option'); 
    opt.value="none";
    opt.id="none";
    opt.selected=true;
    opt.disabled=true;
    opt.hidden=true;
    opt.innerHTML="Select Sheet";
    slt.appendChild(opt);

    for(var i=0;i<wsNo;i++)
    {
        var opt1=document.createElement('option'); 
        opt1.value=ws[i];
        opt1.id=ws[i];
        
        opt1.innerHTML=ws[i];
        slt.appendChild(opt1);
    }
    
    
    for(let i=0;i<wsNo;i++)
    {
        let worksheet=workbook.Sheets[ws[i]];
        let exceljson=XLSX.utils.sheet_to_json(worksheet);
        localStorage.setItem(ws[i],JSON.stringify(exceljson));
    }

    let optio=document.getElementById('list');
    optio.appendChild(slt);

    slt.addEventListener("change",function(){
        let sheet=this.value;
        if(sheet!="none"){
            print_table(sheet);
        }
    },false);

}


let print_table=(sheet)=>{
    if(sheet==undefined)
    {
        let divi=document.getElementById('show');
        divi.innerHTML="Please Enter a File";
        return false;
    }
    var exceljson=JSON.parse(localStorage.getItem(sheet));
    let Attri=Object.keys(exceljson[0]);

    let table=document.createElement('table');
    table.border=1;
    table.className="table table-striped";
    table.style="width:100%";
    let row=table.insertRow(-1);

    for(let i=0;i<Attri.length;i++)
    {
        var header=document.createElement('TH');
        header.innerHTML=Attri[i];
        row.appendChild(header);
    }

    
     for(let i=0;i<exceljson.length;i++)
     {
        var row1=table.insertRow(-1);
        for(let j=0;j<Attri.length;j++)
        {
            var cell1=row1.insertCell(-1);
            var txt1=document.createElement('input');
            txt1.type='text';
            txt1.style="width:100%;text-align:center;";
            txt1.id=Attri[j]+i;
            txt1.value=exceljson[i][Attri[j]];
            txt1.oninput= function(){document.getElementById("chn"+i).disabled=false;};
            cell1.appendChild(txt1);
            row1.appendChild(cell1);
            
        }
        let btn1=document.createElement('input');
         btn1.type="button";
         btn1.value="Change";
         btn1.id="chn"+i;
         btn1.style="width:auto;";
         btn1.disabled=true;
        btn1.addEventListener('click',function(){
            for(let k=0;k<Attri.length;k++){
                exceljson[i][Attri[k]]=document.getElementById(Attri[k]+i).value;
            }
            localStorage.setItem(sheet,JSON.stringify(exceljson));
            this.disabled=true;
        },false);

        let btn2=document.createElement('input');
         btn2.type="button";
         btn2.value="Delete";
         btn2.id="del"+i;
         btn2.style="width:auto;";
         btn2.addEventListener('click',function(){
            exceljson.splice(i,1);

            if(exceljson.length==0)
            {
                localStorage.removeItem(sheet);
                let ws=JSON.parse(localStorage.getItem("sheetNames"));
                
                exceljson=undefined;
                let loc=0;
                while(loc<ws.length)
                {
                    if(ws[loc]===sheet)
                    break;
                    loc++;
                }
                
                ws.splice(loc,1);
                document.getElementById(sheet).remove();
                document.getElementById('none').selected=true;
                if(ws.length==0)
                document.getElementById('select1').remove();
                sheet=undefined;
                if(ws.length!=0)
                localStorage.setItem('sheetNames',JSON.stringify(ws));
                else
                localStorage.removeItem('sheetNames');
                
            }
            else
            localStorage.setItem(sheet,JSON.stringify(exceljson));

            print_table(sheet);
        });


         row1.appendChild(btn1);
         row1.appendChild(btn2);
         

     }

     document.getElementById('show').innerText="";
    let divi=document.getElementById('show');
    divi.appendChild(table);

}


let printing=()=>{
    let excel=XLSX.utils.book_new();
    let sh=JSON.parse(localStorage.getItem('sheetNames'));
    if(sh==undefined)
    {
        alert("File Missing Error..Printing Failed...");
        return ;
    }
    for(let i=0;i<sh.length;i++){
        let js=JSON.parse(localStorage.getItem(sh[i]));
        let ws = XLSX.utils.json_to_sheet(js);
        XLSX.utils.book_append_sheet(excel,ws,sh[i]);
    }
    XLSX.writeFile(excel, "sheetjs.xlsx");

}