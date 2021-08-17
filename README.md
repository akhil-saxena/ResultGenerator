## README

Run the below command to generate results from all `*.xlsx` files in data directory.

`node index.js $(find ./data -type f -name "*.xlsx" -printf "%f\n")`