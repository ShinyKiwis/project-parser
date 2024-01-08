const mammoth = require("mammoth");

filepath = "./project.docx";
function currentTime() {
  return Date.now() / 1000;
}

function extractProjectTitle(inputText) {
  const startString = "Tên đề tài:";
  const endString = "CBHD1";
  const regexPattern = new RegExp(`${startString}([\\s\\S]*?)${endString}`);
  const match = inputText.match(regexPattern);
  let extractedText = match && match[1] ? match[1].trim() : null;
  extractedText = extractedText
    .replace("\n", "")
    .replace("Tiếng Việt:", "")
    .replace("Tiếng Anh", "");
  extractedText = extractedText.split(":").map((text) => {
    return text.trim();
  });

  result = {
    vietnameseTitle: extractedText[0],
    englishTitle: extractedText[1],
  };
  // console.log(result);
  return result;
}

function extractProjectInstructors(text) {
  const emailStartDelimiter = "Email1: ";
  const cbhdStartDelimiter = "CBHD1: ";

  const emailStartIndex = text.indexOf(emailStartDelimiter);
  const cbhdStartIndex = text.indexOf(cbhdStartDelimiter);

  let emailEndIndex = text.indexOf(
    "\n",
    emailStartIndex + emailStartDelimiter.length
  );
  if (emailEndIndex === -1) {
    emailEndIndex = text.length;
  }

  let cbhdEndIndex = text.indexOf(
    "\n",
    cbhdStartIndex + cbhdStartDelimiter.length
  );
  if (cbhdEndIndex === -1) {
    cbhdEndIndex = text.length;
  }

  const emailContent = text
    .substring(emailStartIndex + emailStartDelimiter.length, emailEndIndex)
    .trim();
  const cbhdContent = text
    .substring(cbhdStartIndex + cbhdStartDelimiter.length, cbhdEndIndex)
    .trim();

  // Log the extracted content
  let resultArray = cbhdContent
    .replace(/(CBHD\d*)/g, "")
    .split(":")
    .map((item) => item.trim());
  const instructors = {};
  for (let i = 0; i < resultArray.length; i++) {
    instructors[`CBHD${i + 1}:`] = resultArray[i];
  }
  resultArray = emailContent
    .replace(/(Email\d*)/g, "")
    .split(":")
    .map((item) => item.trim());
  const emails = {};
  for (let i = 0; i < resultArray.length; i++) {
    emails[`Email${i + 1}:`] = resultArray[i];
  }

  // console.log(result);
  return {
    instructors: instructors,
    emails: emails,
  };
}

function extractMajor(inputText) {
  let regexPattern = /Ngành:[^\n]*/;
  let match = inputText.match(regexPattern);
  let extractedText = match ? match[0].trim() : null;
  // Match the strings after ✔ and before ☐ or endline character
  regexPattern = /✔\s*([^☐\n]+)/g;
  extractedText = extractedText.match(regexPattern)[0].replace("✔", "").trim();

  // console.log(extractedText);
  return {
    major: extractedText,
  };
}

function extractBranch(inputText) {
  let regexPattern = /Chương trình đào tạo:[^\n]*/;
  let match = inputText.match(regexPattern);
  let extractedText = match ? match[0].trim() : null;
  // Match the strings after ✔ and before ☐ or endline character
  regexPattern = /✔\s*([^☐\n]+)/g;
  extractedText = extractedText.match(regexPattern)[0].replace("✔", "").trim();

  // console.log(extractedText);
  return {
    branch: extractedText,
  };
}

function extractNumberOfParticipants(inputText) {
  const regexPattern = /Số lượng sinh viên thực hiện:\s*(\d+)/;
  const match = inputText.match(regexPattern);

  // Extract the matched number
  const numberOfParticipants = match ? parseInt(match[1], 10) : null;
  // console.log(numberOfParticipants);
  return {
    numberOfParticipants: numberOfParticipants,
  };
}

function extractParticipants(inputText) {
  const regexPattern = /([^\n]+)\s*-\s*(\d+)(?=[\s\S]*Description)/g;
  const matches = [...inputText.matchAll(regexPattern)];

  const studentInfo = matches.map((match) => ({
    name: match[1].trim(),
    studentID: match[2],
  }));

  // console.log(studentInfo);
  return {
    pariticipants: studentInfo,
  };
}

function extractInfo(inputText, type, startString, endString) {
  const regexPattern = new RegExp(`${startString}([\\s\\S]*?)${endString}`);
  const match = inputText.match(regexPattern);
  const extractedText = match && match[1] ? match[1].trim() : null;

  // console.log(extractedText);
  return {
    [type]: extractedText,
  };
}

function extractRefs(inputText) {
  const startString = "References:";
  const regexPattern = new RegExp(`${startString}([\\s\\S]*)`);
  const match = inputText.match(regexPattern);
  const extractedText = match && match[1] ? match[1].trim() : null;

  return {
    references: extractedText
  }
}

mammoth
  .extractRawText({ path: filepath })
  .then((result) => {
    let text = result.value;
    let startTime = currentTime();
    let project = {};
    project = {...project, ...extractProjectTitle(text)};
    project = {...project, ...extractProjectInstructors(text)};
    project = {...project, ...extractMajor(text)};
    project = {...project, ...extractBranch(text)};
    project = {...project, ...extractNumberOfParticipants(text)};
    project = {...project, ...extractParticipants(text)};
    project = {...project, ...extractInfo(text, "description", "Description:", "Task/Mission")};
    project = {...project, ...extractInfo(text, "task", "Task/Mission", "References")};
    project = {...project, ...extractRefs(text)};
    console.log(currentTime() - startTime);
    console.log(project)
  })
  .catch((error) => {
    console.error(error);
  });
