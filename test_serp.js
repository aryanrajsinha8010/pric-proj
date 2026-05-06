const axios = require('axios');

async function test() {
  try {
    const response = await axios.get('https://serpapi.com/search.json', {
      params: {
        engine: 'google_shopping',
        q: 'iphone',
        api_key: '77a8f890804ba3c501dbb61f752daa3f76188edc6c93add8a8e7b092636c8933',
        hl: 'en',
        gl: 'in'
      }
    });
    console.log(response.data);
  } catch (err) {
    if (err.response) {
      console.error(err.response.data);
    } else {
      console.error(err.message);
    }
  }
}
test();
