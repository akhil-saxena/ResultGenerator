const generator = require('./generator');

const args = process.argv.slice(2);
console.log("Processing files: ", args);

args.forEach((FILE) => {
    generator.run(FILE);
})
