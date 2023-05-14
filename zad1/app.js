const express = require('express');
const app = express();
app.use(express.json());
const fs = require('fs');

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

app.listen(4000, () => {
    console.log("listening to port 4000");
});