const express = require('express');
const app = express();
app.use(express.json());
const fs = require('fs');
const Excel = require('exceljs');

const data = fs.readFileSync('data.json');
const jsonData = JSON.parse(data);

app.get('/',(req,res) => {
    res.send("Welcome to app");
}); 

//GET user by userID
app.get('/users/:id', (req, res) => {
    const userId = parseInt(req.params.id);
    const user = jsonData.users.find(user => user.id === userId);
  
    if (!user) {
      res.status(404).send('User not found');
    } 
    else {
      res.send(user);
    }
});

//GET post by postID
app.get('/posts/:id', (req, res) => {
    const postId = parseInt(req.params.id);
    const post = jsonData.posts.find(post => post.id === postId);
  
    if (!post) {
      res.status(404).send('Post not found');
    } 
    else {
      res.send(post);
    }
  });
  
//Filter by date
app.get('/posts/:dateFrom/:dateTo', (req, res) => {
    const dateFrom = new Date(req.params.dateFrom);
    const dateTo = new Date(req.params.dateTo);
    const postsInRange = jsonData.posts.filter(post => {
        const postDate = new Date(post.last_update);
        return postDate >= dateFrom && postDate <= dateTo;
    });

    if (postsInRange.length === 0) {
        res.status(404).send('No posts found in specified date range');
    } 
    else {
        res.send(postsInRange);
    }
});

// Update email by userID
app.put('/users/update/:id', (req, res) => {
  const userId = parseInt(req.params.id);
  const newEmail = req.body.email;
  const userIndex = jsonData.users.findIndex(user => user.id === userId);

  if (userIndex === -1) {
    res.status(404).send('User not found');
    return;
  }

  jsonData.users[userIndex].email = newEmail;

  fs.writeFile('data.json', JSON.stringify(jsonData, null, 2), (err) => {
    if (err) {
      console.error(err);
      res.status(500).send('Internal server error');
      return;
    }

    res.status(200).send(`User with ID ${userId} email updated to ${newEmail}`);
  });
});

// Add new post
app.post('/posts/add/', (req, res) => {
  const userId = parseInt(req.body.user_id);
  const title = req.body.title;
  const body = req.body.body;

  const user = jsonData.users.find(user => user.id === userId);

  if (!user) {
    res.status(404).send('User not found');
    return;
  }

  const newPost = {
    id: jsonData.posts.length + 1,
    title,
    body,
    user_id: userId,
    last_update: new Date().toISOString()
  };

  jsonData.posts.push(newPost);

  fs.writeFile('data.json', JSON.stringify(jsonData, null, 2), (err) => {
    if (err) {
      console.error(err);
      res.status(500).send('Internal server error');
      return;
    }

    res.status(200).send(`New post added: ${JSON.stringify(newPost)}`);
  });
});

// Generate excel file
app.post('/generate', (req, res) => {
  generateExcelFile()
  .then(() => {
    res.send('Excel file created successfully.');
  })
  .catch((error) => {
    console.error('Error:', error);
  });
});

function excelToJson(workbook){
  const excelData = [];
  let excelTitles = [];
  workbook.worksheets[0].eachRow((row, rowNumber) => {
      if (rowNumber > 0) {
          let rowValues = row.values;
          rowValues.shift();
          if (rowNumber === 1) 
            excelTitles = rowValues;
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
  });
  return excelData;
};

async function generateExcelFile() {
  const fileName = 'nalog.xlsx';

  const wbData = new Excel.Workbook();
  await wbData.xlsx.readFile('data.xlsx');
  const wsData = wbData.getWorksheet('List1');
  const excelData = excelToJson(wbData);

  const wb = new Excel.Workbook();
  const ws = wb.addWorksheet('My Sheet');

  // Set page properties to A4 size
  ws.pageSetup.paperSize = 9;
  ws.pageSetup.orientation = 'landscape'; 

  ws.columns = [
      { width: 6.67 },
      { width: 19.11 },
      { width: 21.89 },
      { width: 21.89 },
      { width: 6.89 },
      { width: 8.56 },
      { width: 7.78 },
      { width: 10.89 },
      { width: 10.67 },
      { width: 10.89 },
  ]; 

  ws.mergeCells('A6:I11');

  const image = wb.addImage({
    filename: 'logo.png',
    extension: 'png',
  });
  ws.addImage(image, {
    tl: { col: 0, row: 0 },
    br: { col: 2, row: 4.5 }
  });

  ws.mergeCells('A5:C5');

  ws.getCell('A5').value = {
    'richText': [
      {'text': 'Predmet: '},
      {'font': {'color': {argb:'FF0000'}},'text': excelData[1].PredmetNaziv + ' ' + excelData[1].PredmetKratica},
  ]};

  const A6 = ws.getCell('A6')
  A6.alignment = {
    vertical: 'middle', 
    horizontal: 'left',
    wrapText: true,
  };
  A6.value = {
    'richText': [
      {'font': {'size': 14, 'bold': true}, 'text': '                                                                                   NALOG ZA ISPLATU\n'},
      {'text': 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.'},
    ]
  };

  // First table
  ws.mergeCells('A12:B12');
  ws.mergeCells('H12:I12');
  ws.mergeCells('A13:B13');
  ws.mergeCells('H13:I13');

  const A12 = ws.getCell('A12');
  A12.value = 'Katedra';
  A12.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  A12.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const C12 = ws.getCell('C12');
  C12.value = 'Studij';
  C12.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  C12.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const D12 = ws.getCell('D12');
  D12.value = 'ak. god.';
  D12.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  D12.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const E12 = ws.getCell('E12');
  E12.value = 'stud. god.';
  E12.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  E12.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const F12 = ws.getCell('F12');
  F12.value = 'početak turnusa';
  F12.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  F12.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const G12 = ws.getCell('G12');
  G12.value = 'kraj turnusa';
  G12.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  G12.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const H12 = ws.getCell('H12');
  H12.value = 'br sati predviđen programom';
  H12.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  H12.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  A13 = ws.getCell('A13')
  A13.value = {
    'richText': [
      {'font': {'color': {argb:'FF0000'}},'text': excelData[1].Katedra},
  ]};
  A13.border = {
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'thin', color: {argb:'00000000'}}
  };
  A13.alignment = {
    horizontal: 'center'
  };

  C13 = ws.getCell('C13')
  C13.value = {
    'richText': [
      {'font': {'color': {argb:'FF0000'}},'text': excelData[1].Studij},
  ]};
  C13.border = {
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'thin', color: {argb:'00000000'}}
  };

  D13 = ws.getCell('D13')
  D13.value = {
    'richText': [
      {'font': {'color': {argb:'FF0000'}},'text': excelData[1].SkolskaGodinaNaziv},
  ]};
  D13.border = {
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'thin', color: {argb:'00000000'}}
  };

  ws.getCell('E13').border = {
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'thin', color: {argb:'00000000'}}
  };
  
  ws.getCell('F13').border = {
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'thin', color: {argb:'00000000'}}
  };
  ws.getCell('G13').border = {
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'thin', color: {argb:'00000000'}}
  };
  
  let P = 0;
  excelData.forEach(e => {
    if(e.PlaniraniSatiPredavanja != '')
    P += parseInt(e.PlaniraniSatiPredavanja)
  })

  let S = 0;
  excelData.forEach(e => {
    if(e.PlaniraniSatiSeminari != '')
    S += parseInt(e.PlaniraniSatiSeminari)
  })

  let V = 0;
  excelData.forEach(e => {
    if(e.PlaniraniSatiVjezbe != '')
    V += parseInt(e.PlaniraniSatiVjezbe)
  })

  H13 = ws.getCell('H13')
  H13.value = {
    'richText': [
      {'text': 'P:'},
      {'font': {'color': {argb:'FF0000'}}, 'text': parseFloat(P)},
      {'text': ' S:'},
      {'font': {'color': {argb:'FF0000'}}, 'text': parseFloat(S)},
      {'text': ' V:'},
      {'font': {'color': {argb:'FF0000'}}, 'text': parseFloat(V)},
  ]};
  H13.border = {
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'thin', color: {argb:'00000000'}}
  };
  H13.alignment = {
    horizontal: 'center'
  }

  ws.getRow(12).height = 60;

  ws.getRow(12).font = {
    bold: true,
  }

  ws.getRow(12).alignment = {
    vertical: 'middle', 
    horizontal: 'center',
    wrapText: true
  } 

  // Second table
  ws.mergeCells('A15:A16');
  ws.mergeCells('B15:B16');
  ws.mergeCells('C15:C16');
  ws.mergeCells('D15:D16');
  ws.mergeCells('H15:H16');
  ws.mergeCells('I15:I16');
  ws.mergeCells('J15:J16');
  ws.mergeCells('N15:N16');
  ws.mergeCells('E15:G15');
  ws.mergeCells('K15:M15');

  ws.getCell('E15').value = 'Sati nastave';
  ws.getCell('K15').value = 'Bruto iznos';

  ws.getRow(15).alignment = {
    vertical: 'middle', 
    horizontal: 'center',
    wrapText: true
  }

  ws.getRow(15).font = {
    bold: true
  }

  ws.getCell('E15').border = {
    top: {style:'medium', color: {argb:'00000000'}},
  };
  ws.getCell('K15').border = {
    top: {style:'medium', color: {argb:'00000000'}},
  };

  ws.getCell('E15').fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };
  ws.getCell('K15').fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  ws.getRow(15).alignment = {
    vertical: 'middle', 
    horizontal: 'center'
  }

  ws.getRow(16).height = 90;

  ws.getRow(16).font = {
    bold: true,
  }
  ws.getRow(16).alignment = {
    vertical: 'middle', 
    horizontal: 'center',
    wrapText: true
  }

  const E15 = ws.getCell('E15');
  E15.value = 'Sati nastave'
  E15.border = {
    top: {style:'medium', color: {argb:'00000000'}}
  };
  E15.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const K15 = ws.getCell('K15');
  K15.value = 'Bruto iznos'
  K15.border = {
    top: {style:'medium', color: {argb:'00000000'}}
  };
  K15.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const A16 = ws.getCell('A16');
  A16.value = 'Redni broj'
  A16.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  A16.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const B16 = ws.getCell('B16');
  B16.value = 'Nastavnik/Suradnik'
  B16.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  B16.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const C16 = ws.getCell('C16');
  C16.value = 'Zvanje'
  C16.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  C16.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const D16 = ws.getCell('D16');
  D16.value = 'Status'
  D16.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  D16.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const E16 = ws.getCell('E16');
  E16.value = 'pred'
  E16.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  E16.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const F16 = ws.getCell('F16');
  F16.value = 'sem'
  F16.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  F16.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const G16 = ws.getCell('G16');
  G16.value = 'vjež'
  G16.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  G16.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const H16 = ws.getCell('H16');
  H16.value = 'Bruto satnica predavanja (EUR)'
  H16.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  H16.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const I16 = ws.getCell('I16');
  I16.value = 'Bruto satnica seminari (EUR)'
  I16.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  I16.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const J16 = ws.getCell('J16');
  J16.value = 'Bruto satnica vježbe (EUR)'
  J16.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  J16.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const K16 = ws.getCell('K16');
  K16.value = 'pred'
  K16.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  K16.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const L16 = ws.getCell('L16');
  L16.value = 'sem'
  L16.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  L16.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const M16 = ws.getCell('M16');
  M16.value = 'vjež'
  M16.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  M16.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  const N16 = ws.getCell('N16');
  N16.value = 'Ukupno za isplatu (EUR)'
  N16.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  N16.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E8E8E8'},
  };

  let ukupnoPred = 0;
  let ukupnoSem = 0;
  let ukupnoVjez = 0;
  for(i = 0; i < excelData.length; i++) {
    rowNum = 17 + i;
    cell0 = ws.getCell('A' + String(rowNum));
    cell0.value = i + 1;
    cell0.alignment = {
      horizontal: 'center',
      vertical: 'middle'
    };
    cell0.border = {
      bottom: {style:'thin', color: {argb:'00000000'}}
    };

    cell1 = ws.getCell('B' + String(rowNum));
    cell1.value = {
      'richText': [
        {'font': {'color': {argb:'FF0000'}},'text': excelData[i].NastavnikSuradnikNaziv},
    ]};
    cell1.border = {
      left: {style:'thin', color: {argb:'00000000'}},
      bottom: {style:'thin', color: {argb:'00000000'}},
      right: {style:'thin', color: {argb:'00000000'}},
    };
    cell1.alignment = {vertical: 'middle'};

    cell2 = ws.getCell('C' + String(rowNum));
    cell2.value = {
      'richText': [
        {'font': {'color': {argb:'FF0000'}},'text': excelData[i].ZvanjeNaziv},
    ]};
    cell2.border = {
      bottom: {style:'thin', color: {argb:'00000000'}},
      right: {style:'thin', color: {argb:'00000000'}},
    };
    cell2.alignment = {
      wrapText: true,
      vertical: 'middle'
    }

    cell3 = ws.getCell('D' + String(rowNum));
    cell3.value = {
      'richText': [
        {'font': {'color': {argb:'FF0000'}},'text': excelData[i].NazivNastavnikStatus},
    ]};
    cell3.border = {
      bottom: {style:'thin', color: {argb:'00000000'}},
      right: {style:'thin', color: {argb:'00000000'}},
    };
    cell3.alignment = {vertical: 'middle'};

    let RealiziraniSatiPredavanja;
    if (excelData[i].RealiziraniSatiPredavanja === '') 
      RealiziraniSatiPredavanja = 0
    else
      RealiziraniSatiPredavanja = parseInt(excelData[i].RealiziraniSatiPredavanja)
     
    cell4 = ws.getCell('E' + String(rowNum));
    cell4.value = RealiziraniSatiPredavanja;
    cell4.font = { color: { argb: 'FF0000' } };
    cell4.border = {
      bottom: {style:'thin', color: {argb:'00000000'}},
      right: {style:'thin', color: {argb:'00000000'}},
    };
    cell4.alignment = {
      horizontal: 'center',
      vertical: 'middle'
    }

    let RealiziraniSatiSeminari;
    if (excelData[i].RealiziraniSatiSeminari === '') 
      RealiziraniSatiSeminari = 0;
    else
      RealiziraniSatiSeminari = parseInt(excelData[i].RealiziraniSatiSeminari);

    cell5 = ws.getCell('F' + String(rowNum));
    cell5.value = RealiziraniSatiSeminari;
    cell5.font = { color: { argb: 'FF0000' } };
    cell5.border = {
      bottom: {style:'thin', color: {argb:'00000000'}},
      right: {style:'thin', color: {argb:'00000000'}},
    };
    cell5.alignment = {
      horizontal: 'center',
      vertical: 'middle'
    }

    let RealiziraniSatiVjezbe;
    if (excelData[i].RealiziraniSatiVjezbe === '') 
      RealiziraniSatiVjezbe = 0;
    else
      RealiziraniSatiVjezbe = parseInt(excelData[i].RealiziraniSatiVjezbe);

    cell6 = ws.getCell('G' + String(rowNum));
    cell6.value = RealiziraniSatiVjezbe;
    cell6.font = { color: { argb: 'FF0000' } };
    cell6.border = {
      bottom: {style:'thin', color: {argb:'00000000'}},
      right: {style:'thin', color: {argb:'00000000'}},
    };
    cell6.alignment = {
      horizontal: 'center',
      vertical: 'middle'
    }

    ws.getCell('H' + String(17 + i)).border = {
      bottom: {style:'thin', color: {argb:'00000000'}},
      right: {style:'thin', color: {argb:'00000000'}}
    }

    ws.getCell('I' + String(17 + i)).border = {
      bottom: {style:'thin', color: {argb:'00000000'}},
      right: {style:'thin', color: {argb:'00000000'}}
    }

    ws.getCell('J' + String(17 + i)).border = {
      bottom: {style:'thin', color: {argb:'00000000'}},
      right: {style:'thin', color: {argb:'00000000'}}
    }

    ws.getCell('K' + String(17 + i)).border = {
      bottom: {style:'thin', color: {argb:'00000000'}},
      right: {style:'thin', color: {argb:'00000000'}}
    }

    ws.getCell('L' + String(17 + i)).border = {
      bottom: {style:'thin', color: {argb:'00000000'}},
      right: {style:'thin', color: {argb:'00000000'}}
    }

    ws.getCell('M' + String(17 + i)).border = {
      bottom: {style:'thin', color: {argb:'00000000'}},
      right: {style:'thin', color: {argb:'00000000'}}
    }

    ws.getCell('N' + String(17 + i)).border = {
      bottom: {style:'thin', color: {argb:'00000000'}},
      right: {style:'medium', color: {argb:'00000000'}}
    }

    ukupnoPred += RealiziraniSatiPredavanja;
    ukupnoSem += RealiziraniSatiSeminari;
    ukupnoVjez += RealiziraniSatiVjezbe;
  };

  mergeFrom = 'A' + String(17 + excelData.length)
  mergeTo = 'C' + String(17 + excelData.length)
  ws.mergeCells(mergeFrom + ':' + mergeTo)
  const ukupno = ws.getCell('A' + String(17 + excelData.length));
  ukupno.value = 'UKUPNO';
  ukupno.alignment = {
    horizontal: 'center',
    vertical: 'middle'
  };
  ukupno.alignment = {
    horizontal: 'center',
    vertical: 'middle'
  };
  ukupno.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  ukupno.font = {bold: true};

  ws.getCell('D' + String(17 + excelData.length)).border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };

  const cellUkupnoPred = ws.getCell('E' + String(17 + excelData.length))
  cellUkupnoPred.value = ukupnoPred;
  cellUkupnoPred.font = {bold: true, color: { argb: 'FF0000'}};
  cellUkupnoPred.alignment = {horizontal: 'center'};
  cellUkupnoPred.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  const cellUkupnoSem = ws.getCell('F' + String(17 + excelData.length))
  cellUkupnoSem.value = ukupnoSem;
  cellUkupnoSem.font = {bold: true, color: { argb: 'FF0000'}};
  cellUkupnoSem.alignment = {horizontal: 'center'};
  cellUkupnoSem.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  const cellUkupnoVjez = ws.getCell('G' + String(17 + excelData.length))
  cellUkupnoVjez.value = ukupnoVjez;
  cellUkupnoVjez.font = {bold: true, color: { argb: 'FF0000'}};
  cellUkupnoVjez.alignment = {horizontal: 'center'};
  cellUkupnoVjez.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('H' + String(17 + excelData.length)).border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('I' + String(17 + excelData.length)).border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('J' + String(17 + excelData.length)).border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('K' + String(17 + excelData.length)).border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('L' + String(17 + excelData.length)).border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('M' + String(17 + excelData.length)).border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('N' + String(17 + excelData.length)).border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.mergeCells('A' + String(20 + excelData.length) + ':C' + String(21 + excelData.length));
  ws.mergeCells('A' + String(26 + excelData.length) + ':C' + String(27 + excelData.length));
  ws.mergeCells('J' + String(26 + excelData.length) + ':L' + String(27 + excelData.length));

  const prodekanicaZaNastavu = ws.getCell('A' + String(20 + excelData.length));
  prodekanicaZaNastavu.value = 'Prodekanica za nastavu i studentska pitanja\nProf. dr. sc.';
  prodekanicaZaNastavu.alignment = {wrapText: true};

  const prodekanZaFinancije = ws.getCell('A' + String(26 + excelData.length));
  prodekanZaFinancije.value = 'Prodekan za financije i upravljanje\nProf. dr. sc.';
  prodekanZaFinancije.alignment = {wrapText: true};

  const dekan = ws.getCell('J' + String(26 + excelData.length));
  dekan.value = 'Dekan\nProf. dr. sc.';
  dekan.alignment = {wrapText: true};

  wb.xlsx
    .writeFile(fileName)
    .then(() => {
      console.log('file created');
    })
    .catch(err => {
      console.log(err.message);
    });
};

app.listen(4000, () => {
    console.log("listening to port 4000");
});