async function test() {
  try {
    const url = new URL('https://serpapi.com/search.json');
    url.searchParams.append('engine', 'google_shopping');
    url.searchParams.append('q', 'tv');
    url.searchParams.append('api_key', '77a8f890804ba3c501dbb61f752daa3f76188edc6c93add8a8e7b092636c8933');
    url.searchParams.append('hl', 'en');
    url.searchParams.append('gl', 'in');

    const response = await fetch(url.toString());
    const data = await response.json();
    console.log("Keys:", Object.keys(data));
    if (data.error) {
      console.log("Error:", data.error);
    }
    console.log("Shopping results count:", data.shopping_results ? data.shopping_results.length : 0);
  } catch (err) {
    console.error(err);
  }
}
test();
