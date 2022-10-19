//Required libraries
const fs=require('fs');
const { degrees } = require('pdf-lib');
const pdf_lib = require ('pdf-lib');
const xlsx=require('xlsx');
const prompt = require('prompt-sync')(); 


//Input and output directories
let exam_directory="./Final_Marked_Directory";
let output_directory='./Final_Totaled_Directory';

//Import mark list
const marks_list=xlsx.readFile('final_marks_list.xlsx');
const cell_range_start=3;
const cell_range_end=73;
const second_sheet_name=marks_list.SheetNames[1];
var worksheet=marks_list.Sheets[second_sheet_name];
var names=[];

//Obtain names from spreadsheet and unpack into array
for (i=cell_range_start;i<=cell_range_end;i++){
    names[i]=worksheet[`D${i}`].v;
}

//Run async function to do the work
write_marks(exam_directory,output_directory,worksheet,names);






//----------------------------------------- Function definition ---------------------------------------------------------------------------------------
async function write_marks(exam_directory,output_directory,worksheet,names){

//Read file names from directory
const files = fs.readdirSync(exam_directory,(err)=>{console.log(err); return;});
//Read through all files
for (const file of files){
    //Obtain file_path

    //Split the file name to find student's first and last name
    var split_strings = file.split("_");
    var name = split_strings[0];
    var split_name = name.split(" ");
    var first_name;
    var last_name;

    //If the student has complex first and last names (that involves spaces), correct by hand, otherwise, assume the first part is first name, second part is last name, which is how eClass downloaded files work
    if (split_name.length==2){
        first_name = split_name[0];
        last_name = split_name[1];
    }
    else{
        console.log(`Unknown first and last name, please enter them, the name on file is ${file}`);
        first_name=prompt('What is the First name?');
        last_name=prompt('What is the last name?');
    }
    //console.log(`First name: ${first_name}, Last name: ${last_name}`);
    //Generate search string for class list
    name_search_string=`${last_name}, ${first_name}`;
    const find_function=(element)=>{return element==name_search_string};

    //Find appropriate element in the spreadsheet that corresponds to this student
    let name_cell = names.findIndex(find_function);
    //If the student could not be found, error and exit, this should never be called theoretically
    if (name_cell==-1){
        console.log(`Could not find this person's marks, their name was ${name_search_string}`);
        return;
    }

    //Otherwise, proceed with grabbing the marks from the spreadsheet
    let Q1_total = worksheet[`AU${name_cell}`].v;
    let Q2_total = worksheet[`AV${name_cell}`].v;
    let Q3_total = worksheet[`AW${name_cell}`].v;
    let Exam_total = worksheet[`AX${name_cell}`].v;
    //console.log(`Name:${name_search_string},Q1: ${Q1_total}, Q2: ${Q2_total},Q3: ${Q3_total},Total: ${Exam_total}`);

    //Write marks on the first page of the student's exam
    await add_marks(exam_directory,file,Q1_total,Q2_total,Q3_total,Exam_total,output_directory);
    console.log(`Completed Totaling for ${name_search_string}`);

    //console.log(name_search_string);

}

}

async function add_marks(input_dir,input_file_name,Q1_mark,Q2_mark,Q3_mark,Total_mark,output_directory){
    var path=`${input_dir}/${input_file_name}`;
    const pdf_bytes=fs.readFileSync(path);
    const doc= await pdf_lib.PDFDocument.load(pdf_bytes);
    const HelveticaFont= await doc.embedFont(pdf_lib.StandardFonts.Helvetica);
    doc.insertPage(0);

    var pages=doc.getPages();
    var font_size=40;
    var front_page=pages[0];
    const{width, height} = front_page.getSize();

    var q1_text=`Q1: ${Q1_mark.toFixed(1)}`;
    var q2_text=`Q2: ${Q2_mark.toFixed(1)}`;
    var q3_text=`Q3: ${Q3_mark.toFixed(1)}`;
    var total_text=`TOTAL: ${Total_mark.toFixed(1)}`;

    mark_text=` ${q1_text} \n\n ${q2_text} \n\n ${q3_text} \n\n------------------------ \n\n ${total_text}\n\n` 

    front_page.drawText(mark_text,{
        x:0.2*width,
        y:0.75*height,
        size: font_size,
        font: HelveticaFont,
    });





    const pdf_bytes_done = await doc.save();
    fs.writeFileSync(`${output_directory}/${input_file_name}`,pdf_bytes_done);


}