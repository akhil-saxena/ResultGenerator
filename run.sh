for f in ./data/*.xlsx ; do mv -- "$f" "${f// /_}" ; done;
node index.js $(find ./data -type f -name "*.xlsx" -printf "%f\n");