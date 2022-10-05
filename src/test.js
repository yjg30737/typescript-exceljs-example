function colName(num) {
  let letters = ''
  while (num >= 0) {
      letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[num % 26] + letters
      num = Math.floor(num / 26) - 1
  }
  return letters
}

let startTime = performance.now();

// for(n = 1; n <= 125; n++)
//       console.log(colName(n-1));

let endTime = performance.now();

// console.log(endTime-startTime);

let startCol = 2;
let startRow = 2;
let endCol = 4

let headerList = getHeaderList(startCol, startRow, endCol); 

function getHeaderList(startCol, startRow, endCol) {
  let startColIdx = startCol-1
  let headerList = []
  for(n = startColIdx; n < startColIdx+endCol; n++) {
    headerList.push(`${colName(n)}${startRow}`);
  }
}