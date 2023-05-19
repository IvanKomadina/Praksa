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
  const fileName = 'nalog.xlsx';

  const wb = new Excel.Workbook();
  const ws = wb.addWorksheet('My Sheet');

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
  ws.getCell('A5').value = 'Predmet: '

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

  ws.getCell('A13').border = {
    bottom: {style:'medium', color: {argb:'00000000'}}
  };
  ws.getCell('C13').border = {
    bottom: {style:'medium', color: {argb:'00000000'}}
  };
  ws.getCell('D13').border = {
    bottom: {style:'medium', color: {argb:'00000000'}}
  };
  ws.getCell('E13').border = {
    bottom: {style:'medium', color: {argb:'00000000'}}
  };
  ws.getCell('F13').border = {
    bottom: {style:'medium', color: {argb:'00000000'}}
  };
  ws.getCell('G13').border = {
    bottom: {style:'medium', color: {argb:'00000000'}}
  };
  ws.getCell('H13').border = {
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };

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
  ws.mergeCells('A25:C25');

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

  const A17 = ws.getCell('A17');
  A17.value = 1;
  A17.alignment = {
    horizontal: 'center'
  };

  const A18 = ws.getCell('A18');
  A18.value = 2;
  A18.alignment = {
    horizontal: 'center'
  };

  const A19 = ws.getCell('A19');
  A19.value = 3;
  A19.alignment = {
    horizontal: 'center'
  };

  const A20 = ws.getCell('A20');
  A20.value = 4;
  A20.alignment = {
    horizontal: 'center'
  };

  const A21 = ws.getCell('A21');
  A21.value = 5;
  A21.alignment = {
    horizontal: 'center'
  };

  const A22 = ws.getCell('A22');
  A22.value = 6;
  A22.alignment = {
    horizontal: 'center'
  };

  const A23 = ws.getCell('A23');
  A23.value = 7;
  A23.alignment = {
    horizontal: 'center'
  };

  const A24 = ws.getCell('A24');
  A24.value = 8;
  A24.alignment = {
    horizontal: 'center'
  };

  const A25 = ws.getCell('A25');
  A25.value = 'UKUPNO';
  A25.alignment = {
    horizontal: 'center',
    vertical: 'middle'
  };
  A25.alignment = {
    horizontal: 'center',
    vertical: 'middle'
  };
  A25.border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };
  A25.font = {
    bold: true
  }

  ws.getCell('D25').border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('E25').border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('F25').border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('G25').border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('H25').border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('I25').border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('J25').border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('K25').border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('L25').border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('M25').border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('N25').border = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  }

  ws.getCell('N17').border = {
    right: {style:'medium', color: {argb:'00000000'}}
  }
  ws.getCell('N18').border = {
    right: {style:'medium', color: {argb:'00000000'}}
  }
  ws.getCell('N19').border = {
    right: {style:'medium', color: {argb:'00000000'}}
  }
  ws.getCell('N20').border = {
    right: {style:'medium', color: {argb:'00000000'}}
  }
  ws.getCell('N21').border = {
    right: {style:'medium', color: {argb:'00000000'}}
  }
  ws.getCell('N22').border = {
    right: {style:'medium', color: {argb:'00000000'}}
  }
  ws.getCell('N23').border = {
    right: {style:'medium', color: {argb:'00000000'}}
  }
  ws.getCell('N24').border = {
    right: {style:'medium', color: {argb:'00000000'}}
  }


  ws.mergeCells('A28:C29');
  ws.mergeCells('A34:C35');
  ws.mergeCells('J34:L35');

  const A28 = ws.getCell('A28')
  A28.value = 'Prodekanica za nastavu i studentska pitanja\nProf. dr. sc.'
  A28.alignment = {
    wrapText: true
  }

  const A34 = ws.getCell('A34')
  A34.value = 'Prodekan za financije i upravljanje\nProf. dr. sc.'
  A34.alignment = {
    wrapText: true
  }

  const J34 = ws.getCell('J34')
  J34.value = 'Dekan\nProf. dr. sc.'
  J34.alignment = {
    wrapText: true
  }

  wb.xlsx
    .writeFile(fileName)
    .then(() => {
      console.log('Excel file created');
    })
    .catch(err => {
      console.log('Error generating Excel file');
    });
});

app.listen(4000, () => {
    console.log("listening to port 4000");
});