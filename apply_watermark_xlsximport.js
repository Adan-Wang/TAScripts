var fs=require('fs');
const { degrees } = require('pdf-lib');
var pdf_lib = require ('pdf-lib');
var xlsx=require('xlsx');

let test_document="test.pdf";
let output_directory='./test_output_2';
//var watermark_text = 'ece456';

const class_list=xlsx.readFile('classlist.xlsx');

generate_watermark_exam(test_document,class_list,output_directory).then(()=>{
    console.log('done');
})



/*addWatermark(test_document,watermark_text).then(()=>{

    console.log('Done');

    
})*/


async function generate_watermark_exam(exam_pdf,class_list,output_directory){
    const first_sheet_name = class_list.SheetNames[0];
    const cell_range_start=2;
    const cell_range_end=72;
    var worksheet = class_list.Sheets[first_sheet_name];
    //console.log('Something');
    for(i=cell_range_start;i<=cell_range_end;i++){
        var ccid_cell = worksheet[`C${i}`];
        var ccid=ccid_cell.v;
        var first_name_cell=worksheet[`A${i}`];
        var first_name=first_name_cell.v;
        var last_name_cell=worksheet[`B${i}`];
        var last_name=last_name_cell.v;

        var output_name=`${first_name}_${last_name}_final_exam`

        await addWatermark(exam_pdf,ccid,output_directory,output_name);
    
    }


}

async function addWatermark(path,text,output_directory,output_name){

    const pdf_bytes=fs.readFileSync(path);
    const doc= await pdf_lib.PDFDocument.load(pdf_bytes);
    
    //console.log('Load Sucessful');

    const CourierFont= await doc.embedFont(pdf_lib.StandardFonts.Courier);
    //console.log('Embed Successful');

    const pages=doc.getPages();

    for (var page of pages){
    //const firstpage=pages[0];
    const{width, height} = page.getSize();
    var font_size=-8*text.length+194;
    var x_pos=(-0.04*text.length+0.42)*width;

    page.drawText(text,{
        x:x_pos,
        y:0.15*height,
        size: font_size,
        font: CourierFont,
        rotate: degrees(25),
        opacity: 0.05,
    });

    page.drawText(text,{
        x:x_pos,
        y:0.45*height,
        size: font_size,
        font: CourierFont,
        rotate: degrees(25),
        opacity: 0.05,
    });

    page.drawText(text,{
        x:x_pos,
        y:0.75*height,
        size: font_size,
        font: CourierFont,
        rotate: degrees(25),
        opacity: 0.05,
    });
    //console.log('Watermark Successful');
    }

    const pdf = await doc.save();
    //console.log('Serialize Successful');
    fs.writeFileSync(`${output_directory}/${output_name}.pdf`,pdf);
    //console.log('Final Save Successful');

}
