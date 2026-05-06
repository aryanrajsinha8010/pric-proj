async function test() {
  try {
    const response = await fetch("http://localhost:3001/api/v1/search?q=iphone");
    const data = await response.json();
    console.log("Is product null?", data.product === null);
    if (data.product) {
      console.log("Product listings count:", data.product.listings ? data.product.listings.length : 0);
    }
  } catch (err) {
    console.error(err);
  }
}
test();
