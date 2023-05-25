
import express from 'express';
import bodyParser from 'body-parser';
import cors from 'cors';
import excelJS from 'exceljs';
const app = express();
const port = 3000;


app.use(bodyParser.urlencoded({
  extended: true
}));

app.use(bodyParser.json({type: 'application/json'}));


app.put('/updateEmail/:id', cors(), (req, res)=>{
  let json = JSON.parse(fs.readFileSync('data.json'));

  fs.readFile("data.json", function (err, data) {
      if (err) {
          console.error(err); 
      }

      JSON.parse(data).users.map((item, key) => {
          if ((item.id) === parseInt(req.params.id)){
              item.email = req.body.email;
              json.users[key] = item;
  
              try {
                  fs.writeFileSync("data.json", JSON.stringify(json, null, 2));
                  console.log("Data successfully saved");
                  return res.send(item);
              } catch (error) {
                  console.log("An error has occurred ", error);
              }
          }
      })
  });
  return null;
})

function getCurrentDate(){
  let date_ob = new Date();
  let date = ("0" + date_ob.getDate()).slice(-2);
  let month = ("0" + (date_ob.getMonth() + 1)).slice(-2);
  let year = date_ob.getFullYear();
  let hours = date_ob.getHours();
  let minutes = date_ob.getMinutes();
  let seconds = date_ob.getSeconds();
  let current = year + "-" + month + "-" + date + " " + hours + ":" + minutes + ":" + seconds;

  return current;
}

// PUT = (userID, title, body) => koji dopustava kreiranje novog post-a (parametri bi bili; user ID, title, body)
app.post('/addPost', cors(), (req, res)=>{
  let json = JSON.parse(fs.readFileSync('data.json'));
  let newPost = {"id": parseInt(json.posts.length + 1), "title": req.body.title, "body": req.body.body, "user_id": req.body.userId, "last_update": getCurrentDate()}
  
  json.posts.push(newPost);

  try {
      fs.writeFileSync("data.json", JSON.stringify(json, null, 2));
      console.log("Data successfully saved");
  } catch (error) {
      console.log("An error has occurred ", error);
  }

  return res.send(newPost);
})

app.listen(port, ()=>{
  console.log("Running on port " + port); 
})


// GET = (userID) => po user ID-u da odgovor budu svi podaci o tom user-u
app.get('/user/:id', cors(), (req, res)=>{
  var user = null;

  fs.readFile("data.json", function (err, data) {
      if (err) {
          console.error(err); 
      }

      user = JSON.parse(data).users.filter(function(key) {
          console.log("KEY: ", key);
          return key.id == req.params.id;
      }).reduce(function(obj, key){
           return key;
          
      }, {});

      return res.send(user);
  });
});

// GET = (postID) => po post ID-u da odgovor budu svi podaci o tom post-u
app.get('/post/:id', cors(), (req, res)=>{
  var post = null;
  
  fs.readFile("data.json", function (err, data) {
      if (err) {
          console.error(err); 
      }
      
      post = JSON.parse(data).posts.filter(function(key) {
          return key.id == req.params.id;
      }).reduce(function(obj, key){
           return key;
          
      }, {});

      return res.send(post);
  });
});

function formattedDate(date){
  var nDate = new Date(date);
  var dd = nDate.getDate();
  var mm = nDate.getMonth()+1; 
  var yyyy = nDate.getFullYear();

  if(dd<10) {
      dd = '0'+dd;
  }
  if(mm<10) {
      mm = '0'+mm;
  }

  var nDate = yyyy + '' + mm + '' + dd ;
  return nDate;

}

// GET = (DatumOd, DatumDo) 
app.get('/postByDate/:startDate/:endDate', cors(), (req, res)=>{
  const startDate = req.params.startDate;
  const endDate = req.params.endDate;

  fs.readFile("data.json", function (err, data) {
      if (err) {
          console.error(err); 
      }
      
      var posts = [];
      JSON.parse(data).posts.map((item) => {
          if ((formattedDate(item.last_update) > formattedDate(startDate)) && (formattedDate(item.last_update) < formattedDate(endDate))){
              console.log("ITEMMM: ", item);
              posts.push(item);
          }
      })

      return res.send(posts);
  });
  
});



app.post('/createExcel', cors(), async (req, res)=>{
  let workbook = new excelJS.Workbook();
  await workbook.xlsx.readFile("data.xlsx");
  let worksheet = workbook.getWorksheet("List1");
  const excelData = excelToJson(workbook);
  
  const subs = [];
      excelData.map((item, i) => {
          if (subs.includes(item.PredmetKratica) === false){
              const fileName = item.PredmetKratica + '.xlsx';   
              const wb = new excelJS.Workbook();
              const ws = wb.addWorksheet('Sheet1');
              var p = 0, s = 0, v = 0;
              
              wb.xlsx.writeFile(fileName).then(() => {
                console.log('File created');
            }).catch(err => {
                console.log(err.message);
            });

            ws.columns = [
              { header: '', key: 'A', width: 7 },
              { header: '', key: 'B', width: 19 },
              { header: '', key: 'C', width: 22 },
              { header: '', key: 'D', width: 22 },
              { header: '', key: 'E', width: 7 },
              { header: '', key: 'F', width: 9 },
              { header: '', key: 'G', width: 8 },
              { header: '', key: 'H', width: 11 },
              { header: '', key: 'I', width: 11 },
              { header: '', key: 'J', width: 11 },
              { header: '', key: 'K', width: 9 },
              { header: '', key: 'L', width: 9 },
              { header: '', key: 'M', width: 9 },
              { header: '', key: 'N', width: 9 },
              { header: '', key: 'O', width: 9 },
              { header: '', key: 'P', width: 9 }
            ];

             
              ws.mergeCells('A5:C5');
              ws.getCell('A5').value = {
                  'richText': [
                      {'font': {'color': {argb: '000000'}},'text': 'Predmet: '},
                      {'font': {'color': {argb: 'FF0000'}},'text': item.PredmetKratica + ' ' + item.PredmetNaziv}
                  ]
              }

              ws.mergeCells('A6:I11'); 
              ws.getCell('A6').value = {
                  'richText': [
                      {'font': {'size': 14,'color': {argb: '000000'}, bold: true},'alignment': {'vertical': 'middle', 'horizontal': 'center'},'text': 'NALOG ZA ISPLATU\r\n'},
                      {'font': {'color': {argb: '000000'}},'text': 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.'}
                  ]
              };
              ws.getCell(`A6`).alignment = {
                  vertical: 'middle', horizontal: 'left',
                  wrapText: true
              };

              //Prva tablica
              firstTable(ws);
              
              excelData.map((item) => {
                  p += item.PlaniraniSatiPredavanja != '' ? parseInt(item.PlaniraniSatiPredavanja) : 0;
                  s += item.PlaniraniSatiSeminari != '' ? parseInt(item.PlaniraniSatiSeminari) : 0;
                  v += item.PlaniraniSatiVjezbe != '' ? parseInt(item.PlaniraniSatiVjezbe) : 0;
              })

              ws.mergeCells('H13:I13');
              ws.getCell('H13').value = {
                  'richText': [
                      {'text': 'P:'},
                      {'text': p},
                      {'text': ' S:'},
                      {'text': s},
                      {'text': ' V:'},
                      {'text': v}
                  ]
              };

              ws.getCell('H13').alignment = {
                  vertical: 'middle', horizontal: 'center',
                  wrapText: true
              };

              //Druga tablica

              index = secondTable(ws);


              ws.mergeCells('A' + (index + 3) +':C' + (index + 4));
              ws.getCell('A' + (index + 3)).value = {
                  'richText': [
                      {'text': 'Prodekanica za nastavu i studentska pitanja\r\nProf. dr. sc.'},
                      {'font': {'color': {argb: 'FF0000'}},'text': ' Ime Prezime'}
                  ]
              };
              ws.getCell('A' + (index + 3)).alignment = {
                  vertical: 'middle', horizontal: 'left',
                  wrapText: true
              };

              ws.mergeCells('A' + (index + 9) +':C' + (index + 10));
              ws.getCell('A' + (index + 9)).value = {
                  'richText': [
                      {'text': 'Prodekan za financije i upravljanje\r\nProf. dr. sc.'},
                      {'font': {'color': {argb: 'FF0000'}},'text': ' Ime Prezime'}
                  ]
              };

              ws.getCell(`A` + (index + 9)).alignment = {
                  vertical: 'middle', horizontal: 'left',
                  wrapText: true
              };

              ws.mergeCells('J' + (index + 9) +':L' + (index + 10));
              ws.getCell('J' + (index + 9)).value = {
                  'richText': [
                      {'text': 'Dekan\r\nProf. dr. sc.'},
                      {'font': {'color': {argb: 'FF0000'}},'text': ' Ime Prezime'}
                  ]
              };
              ws.getCell(`J` + (index + 9)).alignment = {
                  vertical: 'middle', horizontal: 'left',
                  wrapText: true
              };

              subs.push(item.PredmetKratica);
              wb.xlsx.writeFile(fileName).then(() => {
                  console.log('File created');
              }).catch(err => {
                  console.log(err.message);
              });
             }
      });
  })

function firstTable(ws){

  ws.mergeCells('A12:B12');
  ws.getCell('A12').value = 'Katedra';
  
  ws.getCell('A12').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };

  ws.getCell('A12').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('A12').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('A12').font = {
      bold: true
  };

  ws.getCell('C12').value = 'Studij';
  ws.getCell('C12').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('C12').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('C12').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('C12').font = {
      bold: true
  };

  ws.getCell('D12').value = 'ak. god.';
  ws.getCell('D12').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('D12').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('D12').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('D12').font = {
      bold: true
  };

  ws.getCell('E12').value = 'stud. god.';
  ws.getCell('E12').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('E12').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('E12').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('E12').font = {
      bold: true
  };

  ws.getCell('F12').value = 'početak turnusa';
  ws.getCell('F12').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('F12').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('F12').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('F12').font = {
      bold: true
  };

  ws.getCell('G12').value = 'kraj turnusa';
  ws.getCell('G12').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('G12').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('G12').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('G12').font = {
      bold: true
  };

  ws.mergeCells('H12:I12');
  ws.getCell('H12').value = 'br sati predviđen programom';
  ws.getCell('H12').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('H12').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('H12').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('H12').font = {
      bold: true
  };

  const row = ws.getRow(12);
  row.height = 75;

  ws.mergeCells('A13:B13');
  ws.getCell('A13').value = item.Katedra;
  ws.getCell('A13').alignment = {
    vertical: 'middle', horizontal: "center",
    wrapText: true
  };
  ws.getCell('A13').border = {
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'thin', color: {argb:'000'}}
  };

  ws.getCell('A13').font = {
      color: { argb: 'FF0000'}
  };

  ws.getCell('C13').value = item.Studij;
  ws.getCell('C13').alignment = {
    vertical: 'middle', horizontal: "left",
    wrapText: true
  };
  ws.getCell('C13').border = {
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'thin', color: {argb:'000'}}
  };

  ws.getCell('C13').font = {
      color: { argb: 'FF0000'}
  };


  ws.getCell('D13').value = item.SkolskaGodinaNaziv;
  ws.getCell('D13').alignment = {
    vertical: 'middle', horizontal: "left",
    wrapText: true
  };
  ws.getCell('D13').border = {
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'thin', color: {argb:'000'}}
  };

  ws.getCell('D13').font = {
      color: { argb: 'FF0000'}
  };

  ws.getCell('E13').value = item.PKSkolskaGodina;
  ws.getCell('E13').alignment = {
    vertical: 'middle', horizontal: "left",
    wrapText: true
  };
  ws.getCell('E13').border = {
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'thin', color: {argb:'000'}}
  };

  ws.getCell('E13').font = {
      color: { argb: 'FF0000'}
  };

  ws.getCell('F13').alignment = {
    vertical: 'middle', horizontal: "left",
    wrapText: true
  };
  ws.getCell('F13').border = {
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'thin', color: {argb:'000'}}
  };

  ws.getCell('F13').font = {
      color: { argb: 'FF0000'}
  };

  ws.getCell('G13').alignment = {
    vertical: 'middle', horizontal: "left",
    wrapText: true
  };
  ws.getCell('G13').border = {
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'thin', color: {argb:'000'}}
  };

  ws.getCell('G13').font = {
      color: { argb: 'FF0000'}
  };

  

}

function secondTable(ws){
  ws.mergeCells('A15:A16');
  ws.getCell('A15').value = 'Redni broj';
  ws.getCell('A15').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('A15').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('A15').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('A15').font = {
      bold: true
  };

  ws.mergeCells('B15:B16');
  ws.getCell('B15').value = 'Nastavnik/Suradnik';
  ws.getCell('B15').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('B12').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('B12').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('B12').font = {
      bold: true
  };

  ws.mergeCells('C15:C16');
  ws.getCell('C15').value = 'Zvanje';
  ws.getCell('C15').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('C15').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('C15').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('C15').font = {
      bold: true
  };

  ws.mergeCells('D15:D16');
  ws.getCell('D15').value = 'Status';
  ws.getCell('D15').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('D15').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('D15').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('D15').font = {
      bold: true
  };

  ws.mergeCells('E15:G15');
  ws.getCell('E15').value = 'Sati nastave';
  ws.getCell('E15').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('E15').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('E15').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('E15').font = {
      bold: true
  };

  ws.getCell('E16').value = 'pred';
  ws.getCell('E16').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('E16').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('E16').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('E16').font = {
      bold: true
  };

  ws.getCell('F16').value = 'sem';
  ws.getCell('F16').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('F16').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('F16').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('F16').font = {
      bold: true
  };

  ws.getCell('G16').value = 'vjež';
  ws.getCell('G16').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('G16').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('G16').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('G16').font = {
      bold: true
  };


  ws.mergeCells('H15:H16');
  ws.getCell('H15').value = 'Bruto satnica predavanja (EUR)';
  ws.getCell('H15').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('H15').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('H15').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('H15').font = {
      bold: true
  };

  ws.mergeCells('I15:I16');
  ws.getCell('I15').value = 'Bruto satnica seminar (EUR)';
  ws.getCell('I15').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('I15').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('I15').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('I15').font = {
      bold: true
  };

  ws.mergeCells('J15:J16');
  ws.getCell('J15').value = 'Bruto satnica vježbe (EUR)';
  ws.getCell('J15').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('J15').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('J15').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('J15').font = {
      bold: true
  };

  ws.mergeCells('K15:M15');
  ws.getCell('K15').value = 'Bruto iznos';
  ws.getCell('K15').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('K15').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('K15').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('K15').font = {
      bold: true
  };

  ws.getCell('K16').value = 'pred';
  ws.getCell('K16').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('K16').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('K16').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('K16').font = {
      bold: true
  };

  ws.getCell('L16').value = 'sem';
  ws.getCell('L16').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('L16').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('L16').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('L16').font = {
      bold: true
  };

  ws.getCell('M16').value = 'vjež';
  ws.getCell('M16').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('M16').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('M16').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('M16').font = {
      bold: true
  };

  ws.mergeCells('N15:N16');
  ws.getCell('N15').value = 'Ukupno za isplatu (EUR)';
  ws.getCell('N15').alignment = {
    vertical: 'middle', horizontal: 'center',
    wrapText: true
  };
  ws.getCell('N15').border = {
    top: {style:'medium', color: {argb:'000'}},
    left: {style:'medium', color: {argb:'000'}},
    bottom: {style:'medium', color: {argb:'000'}},
    right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('N15').fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{argb:'eeeeee'}
  };

  ws.getCell('N15').font = {
      bold: true
  };
  const row2 = ws.getRow(16);
  row2.height = 105;

  
  var index = 17;
  var ukupnoPredavanja = 0, ukupnoSeminari = 0, ukupnoVjezbe = 0;
  excelData.map((it, i) => {
      ws.getCell('A' + index).value = i + 1; // 1.
      getIndexStyle(ws, 'A' + index, i + 1);
      getRowStyle(ws, 'B' + index, "left");
      ws.getCell('B' + index).value = it.NastavnikSuradnikNaziv;
      getRowStyle(ws, 'C' + index, "left");
      ws.getCell('C' + index).value = it.ZvanjeNaziv;
      getRowStyle(ws, 'D' + index, "left");
      ws.getCell('D' + index).value = it.NazivNastavnikStatus;
      getRowStyle(ws, 'E' + index, "center");
      ws.getCell('E' + index).value = it.RealiziraniSatiPredavanja != '' ? it.RealiziraniSatiPredavanja : 0;
      ukupnoPredavanja += it.RealiziraniSatiPredavanja != '' ? parseInt(it.RealiziraniSatiPredavanja) : 0;
      getRowStyle(ws, 'F' + index, "center");
      ws.getCell('F' + index).value = it.RealiziraniSatiSeminari != '' ? it.RealiziraniSatiSeminari : 0;
      ukupnoSeminari += it.RealiziraniSatiSeminari != '' ? parseInt(it.RealiziraniSatiSeminari) : 0;
      getRowStyle(ws, 'G' + index, "center");
      ws.getCell('G' + index).value = it.RealiziraniSatiVjezbe != '' ? it.RealiziraniSatiVjezbe : 0;
      ukupnoVjezbe += it.RealiziraniSatiVjezbe != '' ? parseInt(it.RealiziraniSatiVjezbe) : 0;
      getRowStyle(ws, 'H' + index, "left");
      getRowStyle(ws, 'I' + index, "left");
      getRowStyle(ws, 'J' + index, "left");
      getRowStyle(ws, 'K' + index, "right");
      getRowStyle(ws, 'L' + index, "right");
      getRowStyle(ws, 'M' + index, "right");
      getRowStyle(ws, 'N' + index, "right");
      index++;
  })

  ws.mergeCells('A' + index + ':C' + index);
  ws.getCell('A' + index + ':C' + index).value = 'UKUPNO';
  ws.getCell('A' + index + ':C' + index).alignment = {
    vertical: 'middle', horizontal: "center",
    wrapText: true
  };

  ws.getCell('A' + index + ':C' + index).border = {
      top: {style:'medium', color: {argb:'000'}},
      left: {style:'medium', color: {argb:'000'}},
      bottom: {style:'medium', color: {argb:'000'}},
      right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('A' + index + ':C' + index).font = {
      bold: true
  };

  ws.getCell('D' + index).alignment = {
    vertical: 'middle', horizontal: "center",
    wrapText: true
  };

  ws.getCell('D' + index).border = {
      top: {style:'medium', color: {argb:'000'}},
      left: {style:'medium', color: {argb:'000'}},
      bottom: {style:'medium', color: {argb:'000'}},
      right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('D' + index).font = {
      bold: true
  };

  ws.getCell('E' + index).alignment = {
    vertical: 'middle', horizontal: "center",
    wrapText: true
  };

  ws.getCell('E' + index).border = {
      top: {style:'medium', color: {argb:'000'}},
      left: {style:'medium', color: {argb:'000'}},
      bottom: {style:'medium', color: {argb:'000'}},
      right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('E' + index).font = {
      bold: true,
      color: { argb: 'FF0000'}
  };

  ws.getCell('E' + index).value = ukupnoPredavanja;
  ws.getCell('E' + index).alignment = {
    vertical: 'middle', horizontal: "center",
    wrapText: true
  };

  ws.getCell('E' + index).border = {
      top: {style:'medium', color: {argb:'000'}},
      left: {style:'medium', color: {argb:'000'}},
      bottom: {style:'medium', color: {argb:'000'}},
      right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('E' + index).font = {
      bold: true,
      color: { argb: 'FF0000'}
  };


  ws.getCell('F' + index).value = ukupnoSeminari;
  ws.getCell('F' + index).alignment = {
    vertical: 'middle', horizontal: "center",
    wrapText: true
  };

  ws.getCell('F' + index).border = {
      top: {style:'medium', color: {argb:'000'}},
      left: {style:'medium', color: {argb:'000'}},
      bottom: {style:'medium', color: {argb:'000'}},
      right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('F' + index).font = {
      bold: true,
      color: { argb: 'FF0000'}
  };

  ws.getCell('G' + index).value = ukupnoVjezbe;
  ws.getCell('G' + index).alignment = {
    vertical: 'middle', horizontal: "center",
    wrapText: true
  };

  ws.getCell('G' + index).border = {
      top: {style:'medium', color: {argb:'000'}},
      left: {style:'medium', color: {argb:'000'}},
      bottom: {style:'medium', color: {argb:'000'}},
      right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('G' + index).font = {
      bold: true,
      color: { argb: 'FF0000'}
  };

  ws.getCell('I' + index).alignment = {
    vertical: 'middle', horizontal: "center",
    wrapText: true
  };

  ws.getCell('I' + index).border = {
      top: {style:'medium', color: {argb:'000'}},
      left: {style:'medium', color: {argb:'000'}},
      bottom: {style:'medium', color: {argb:'000'}},
      right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('I' + index).font = {
      bold: true,
      color: { argb: 'FF0000'}
  };

  ws.getCell('J' + index).alignment = {
    vertical: 'middle', horizontal: "center",
    wrapText: true
  };

  ws.getCell('J' + index).border = {
      top: {style:'medium', color: {argb:'000'}},
      left: {style:'medium', color: {argb:'000'}},
      bottom: {style:'medium', color: {argb:'000'}},
      right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('J' + index).font = {
      bold: true,
      color: { argb: 'FF0000'}
  };

  ws.getCell('K' + index).alignment = {
    vertical: 'middle', horizontal: "right",
    wrapText: true
  };

  ws.getCell('K' + index).border = {
      top: {style:'medium', color: {argb:'000'}},
      left: {style:'medium', color: {argb:'000'}},
      bottom: {style:'medium', color: {argb:'000'}},
      right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('K' + index).font = {
      bold: true,
      color: { argb: 'FF0000'}
  };

  ws.getCell('L' + index).alignment = {
    vertical: 'middle', horizontal: "right",
    wrapText: true
  };

  ws.getCell('L' + index).border = {
      top: {style:'medium', color: {argb:'000'}},
      left: {style:'medium', color: {argb:'000'}},
      bottom: {style:'medium', color: {argb:'000'}},
      right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('L' + index).font = {
      bold: true,
      color: { argb: 'FF0000'}
  };

  ws.getCell('M' + index).alignment = {
    vertical: 'middle', horizontal: "right",
    wrapText: true
  };

  ws.getCell('M' + index).border = {
      top: {style:'medium', color: {argb:'000'}},
      left: {style:'medium', color: {argb:'000'}},
      bottom: {style:'medium', color: {argb:'000'}},
      right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('M' + index).font = {
      bold: true,
      color: { argb: 'FF0000'}
  };

  ws.getCell('N' + index).alignment = {
    vertical: 'middle', horizontal: "right",
    wrapText: true
  };

  ws.getCell('N' + index).border = {
      top: {style:'medium', color: {argb:'000'}},
      left: {style:'medium', color: {argb:'000'}},
      bottom: {style:'medium', color: {argb:'000'}},
      right: {style:'medium', color: {argb:'000'}}
  };

  ws.getCell('N' + index).font = {
      bold: true,
      color: { argb: 'FF0000'}
  };

  return index;
}


function getRowStyle(ws, cell, align){
  ws.getCell(cell).alignment = {
      vertical: 'middle', horizontal: align,
      wrapText: true
  };

  ws.getCell(cell).border = {
      top: {style:'thin', color: {argb:'000'}},
      left: {style:'thin', color: {argb:'000'}},
      bottom: {style:'thin', color: {argb:'000'}},
      right: {style:'thin', color: {argb:'000'}}
  };

  ws.getCell(cell).font = {
      color: { argb: 'FF0000'}
  };
}



function getIndexStyle(ws, index, i){
  ws.getCell(index).value = i;

      ws.getCell(index).alignment = {
          vertical: 'middle', horizontal: 'center',
          wrapText: true
      };

      ws.getCell(index).border = {
          top: {style:'thin', color: {argb:'000'}},
          left: {style:'thin', color: {argb:'000'}},
          bottom: {style:'thin', color: {argb:'000'}},
          right: {style:'thin', color: {argb:'000'}}
      };

      ws.getCell(index).font = {
          color: { argb: '000000'}
      }
}



function excelToJson(workbook){
  const excelData = [];
  let excelTitles = [];
  workbook.worksheets[0].eachRow((row, rowNumber) => {
      if (rowNumber > 0) {
          let rowValues = row.values;
          rowValues.shift();
          if (rowNumber === 1) excelTitles = rowValues;
          else {
              let rowObject = {}
              for (let i = 0; i < excelTitles.length; i++) {
                  let title = excelTitles[i];
                  let value = rowValues[i] ? rowValues[i] : '';
                  rowObject[title] = value;
              }
              excelData.push(rowObject);
          }
      }
  })
  return excelData;
}



app.listen(port, ()=>{
  console.log("Running on port " + port); 
})
