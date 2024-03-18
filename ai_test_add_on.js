var doc = DocumentApp.getActiveDocument()
var body = doc.getBody()
function onOpen() {
   DocumentApp.getUi().createMenu("zyBooks AI Add-on")
    .addItem("Create multiple choice questions", "aiMCQ")
    .addItem("Create learning objectives", "aiLOs")
    .addItem("Generate MCQ XML", "mcqXML")
    .addItem("Generate Test Bank XML", "testXML")
    .addItem("Summarize text", "aiText")
    .addToUi();
}

var text_guidelines = `
Authoring guidelines
- Use shorter words. Ex: also instead of additionally, about instead of approximately, but instead of however,
The sentence "Additionally, the numerous requested modifications demonstrate the organization's dissatisfaction." 
should be written as "Also, the many requested changes show the firm's displeasure."
- Do not use unneeded words or phrases. Ex: many instead of a variety of
- Use the posessive form. Ex: cat's eye instead of eye of the cat
- Do not use phrases like "to be clear", "note that", "greatly", "very", "sorted in order", "it is clear",
"one can see", "let us consider"
- Avoid using the word there. Ex: "Mice are hiding in the attic" should be used instead of "There are mice hiding in the attic."
- Use direct sentences. Ex: "A median is not skewed by a few large or small values. Thus, a median may better represent a "typical" value rather than a mean (aka "average" value)." shoud be used instead of "The basic advantage of the median in describing data compared to the mean (often simply described as the "average") is that it is not skewed so much by a small proportion of extremely large or small values, and so it may give a better idea of a "typical" value."
- For more technical topics, key terms should be explicitly defined.
"Authors should explicitly define key terms. Ex: A smart phone is a mobile phone that has functions common with computers, like internet access and downloading of apps. Definitions help learners see what terms will be important to remember in later content. A definition's term is highlighted, and the term plus definition will appear in the zyBook's content explorer (akin to a glossary). Key considerations:

Quantity: Authors should define important terms but avoid too many definitions. Defining too many terms prevents a learner from knowing which are most important. Striking the balance is an art.
Concise and precise: Authors should strive to create definitions that balance conciseness and preciseness. Authors commonly write overly precise and hence long definitions, defensible to experts who may point out incompleteness or flaws with a concise definition, but incomprehensible to learners.
Define in the singular: Authors should use a singular term in a definition whenever possible. A plural term is ambiguous; in "Cars have engines that propel the car," does a car have multiple engines or just one? Better: "A car has an engine that propels the car.""
- Do not use the word it. Ex: "Regular activity improves mental health by easing stress, anxiety, and depression, while lifting self-esteem. It also sharpens cognitive function." should be written as "Regular activity improves mental health by easing stress, anxiety, and depression, while lifting self-esteem. Regular activity also sharpens cognitive function." instead.
- Do not use personal pronouns
- Do not use contractions. Ex: cannot instead of can't
- Use explicit subject. Ex: "Engaging in regular exercise aids in weight management, calorie burning, and muscle mass building, which increases metabolism. This helps prevent chronic diseases such as heart disease, diabetes, and certain cancers." should be written as 
"Engaging in regular exercise aids in weight management, calorie burning, and muscle mass building, which increases metabolism and helps prevent chronic diseases like diabetes."
`
var text_instructions = `Summarize the text using less than 150 words. Follow the authoring guidelines.`

var mcq_xml = `
&lt;zyQSetMultipleChoice caption=&quot;FIX_ME&quot; id=&quot;FIX_ME&quot;&gt;
&lt;zyQ&gt;
&lt;zyQText&gt;Question&lt;/zyQText&gt;
&lt;zyQChoice correct='true'&gt;
&lt;zyQAns&gt;Correct choice&lt;/zyQAns&gt;
&lt;zyQExpl&gt;Explanation&lt;/zyQExpl&gt;
&lt;/zyQChoice&gt;
&lt;zyQChoice&gt;
&lt;zyQAns&gt;Incorrect choice 1&lt;/zyQAns&gt;
&lt;zyQExpl&gt;Explanation&lt;/zyQExpl&gt;
&lt;/zyQChoice&gt;&lt;zyQChoice&gt;
&lt;zyQAns&gt;Incorrect choice 2&lt;/zyQAns&gt;
&lt;zyQExpl&gt;Explanation&lt;/zyQExpl&gt;
&lt;/zyQChoice&gt;
&lt;/zyQ&gt;
&lt;zyQSetMultipleChoice&gt;
`
var mcq_guidelines = `
Authoring Guidelines
- If possible, follow this format when asking questions
Review: Focusing the learner on the key concepts, potentially with further explanation.
Explore: Helping the learner understand specifics of the concepts.
Expand: Teaching even more concepts.
- Authors often design a sequence of learning questions that cover incrementally 
harder aspects of a concept, especially via "explore" learning questions.
- Concise questions
Focusing on one concept
Using fewer words
Using fill-in-the-blank
Using a common instruction
- Concise choices
Pulling out common words into the question
Using fill-in-the-blank
- Explanation guidelines
Using fewer words—as.
Staying focused—wordy explanations often come with unnecessary info.
Making explanations concrete—using values, equations, or similar concrete items as in "2 + 2 = 4," 
rather than textual discussions as in "adding 2 with 2 yields a result of 4" .
- Creating choices
Be short.
avoid "all of the above" or "none of the above" choices.
Be about the same length. Exceptions are common, though.
Be parallel, meaning having the same form.
- Explanations
Explicitly explaining right answers
Explicitly explaining wrong answers
`
var mcq_instructions_xml = `Write multiple choice questions for this text with explanations for both correct and incorrect choices. Do not highlight html syntax or use backticks in the output. The id attribute should have be 8-4-4-4-12 alphanumeric characters like 96104e5d-342a-fe9a-f080-f3a5583ebe42 but should be random each time.
  
Example:
  
&lt;zyQSetMultipleChoice caption=&quot;FIX_ME&quot; id=&quot;FIX_ME&quot;&gt;
&lt;zyQ&gt;
&lt;zyQText&gt;Which of the following benefits is primarily associated with regular physical activity?&lt;/zyQText&gt;
&lt;zyQChoice correct='true'&gt;
&lt;zyQAns&gt;Improved mental health&lt;/zyQAns&gt;
&lt;zyQExpl&gt;Improved mental health. Regular physical activity helps reduce stress, anxiety, and depression while improving mood, self-esteem, and cognitive function through the release of endorphins, which are natural mood boosters.&lt;/zyQExpl&gt;
&lt;/zyQChoice&gt;
&lt;zyQChoice&gt;
&lt;zyQAns&gt;Increased risk of heart disease&lt;/zyQAns&gt;
&lt;zyQExpl&gt;Increased risk of heart disease is incorrect because physical activity actually reduces the risk of heart disease by improving blood circulation, lowering blood pressure, and improving cholesterol levels.&lt;/zyQExpl&gt;
&lt;/zyQChoice&gt;&lt;zyQChoice&gt;
&lt;zyQAns&gt;Decreased muscle mass&lt;/zyQAns&gt;
&lt;zyQExpl&gt;Decreased muscle mass is incorrect as physical activity, especially strength training exercises, helps to build and maintain muscle mass.&lt;/zyQExpl&gt;
&lt;/zyQChoice&gt;
&lt;/zyQ&gt;
&lt;zyQSetMultipleChoice&gt;
`
var test_instructions = `Write multiple choice questions for this text. Do not highlight html syntax or use backticks in the output. The guid attribute should have be 8-4-4-4-12 alphanumeric characters like 96104e5d-342a-fe9a-f080-f3a5583ebe42 but should be random each time.
  
Example:
  
&lt;question xsi:type='MultipleChoiceQuestion' guid='a706db50-7ad5-4a03-b374-5bbdbf40b5b4'&gt; &lt;instructions&gt; Machine learning uses algorithms and models to _____. &lt;/instructions&gt; &lt;choices&gt; &lt;choice&gt;make predictions&lt;/choice&gt; &lt;choice&gt;discover patterns in data&lt;/choice&gt; &lt;choice correct='true'&gt;make predictions and discover patterns in data&lt;/choice&gt; &lt;choice&gt;reformat datasets into dataframes&lt;/choice&gt; &lt;/choices&gt; &lt;/question&gt;
`

var test_xml = `
&lt;question xsi:type='MultipleChoiceQuestion' guid='FIX_ME'&gt; &lt;instructions&gt; A(n) _____ is a mathematical function for describing the relationship between input features and output features. &lt;/instructions&gt; &lt;choices&gt; &lt;choice&gt;algorithm&lt;/choice&gt; &lt;choice correct='true'&gt;model&lt;/choice&gt; &lt;choice&gt;formula&lt;/choice&gt; &lt;choice&gt;prediction&lt;/choice&gt; &lt;/choices&gt; &lt;/question&gt;
`

var mcq_instructions_text = `Write multiple choice questions for this text with explanations for both correct and incorrect choices. 
Add html tags to the output. Do not highlight html syntax or use backticks in the output.
Choices should be an ordered list <ol type="a">. Do not mention the text in the explanations.

Example:

<h3>Question 1</h3>

<ol>Which of the following benefits is primarily associated with regular physical activity?
<ol type="a">
<li>Increased risk of heart disease</li>
<li>Improved mental health</li>
<li>Decreased muscle mass</li>
<li>Lowered energy levels</li>
</ol>

<h4>Explanation for Question 1</h4>

Correct answer: <br/>
b. Improved mental health. Regular physical activity helps reduce stress, anxiety, and depression while improving mood, self-esteem, and cognitive function through the release of endorphins, which are natural mood boosters.<br/><br/>

Incorrect choices:<br/>
a. Increased risk of heart disease is incorrect because physical activity actually reduces the risk of heart disease by improving blood circulation, lowering blood pressure, and improving cholesterol levels.<br/>
c. Decreased muscle mass is incorrect as physical activity, especially strength training exercises, helps to build and maintain muscle mass.<br/>
d. Lowered energy levels are incorrect because regular physical activity increases energy levels and improves overall well-being.
</ol>
`
var lo_instructions = `Write at most 5 learning objectives that follow Bloom's taxonomy for this content. 
Display output as HTML bullet points. Do not identify the Bloom's taxonomy guidelines in the output.

Example:
  
<h3>Learning objectives</h3>
<ul>
<li>Analyze the effects of regular physical activity on maintaining a healthy weight and reducing the risk of chronic diseases.</li>
<li>Compare and contrast the impact of physical activity on cardiovascular health and the development of atherosclerosis.</li>
<li>Evaluate the significance of physical activity in strengthening bones and muscles and its role in preventing osteoporosis and injuries.</li>
<li>Summarize how physical activity contributes to mental health improvement, including reduction of stress, anxiety, and depression.</li>
<li>Create a personalized physical activity plan that incorporates at least 150 minutes of moderate-intensity exercise per week to improve overall health and quality of life.</li>
</ul>
`

function aiMCQ(){
  var htmlOutput = HtmlService.createHtmlOutput('<p>Loading questions...</p>')
    .setTitle('zyBooks AI Add-on')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(htmlOutput);
  var promptBase = mcq_instructions_text + mcq_guidelines;
  aiBlog(promptBase)
}

function aiText(){
  var htmlOutput = HtmlService.createHtmlOutput('<p>Loading summary...</p>')
    .setTitle('zyBooks AI Add-on')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(htmlOutput);
  var promptBase = text_guidelines + text_instructions;
  aiBlog(promptBase)
}

function aiLOs(){
  var htmlOutput = HtmlService.createHtmlOutput('<p>Loading learning objectives...</p>')
    .setTitle('zyBooks AI Add-on')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(htmlOutput);

  var promptBase = lo_instructions
  aiBlog(promptBase)
}

function mcqXML(){
  var htmlOutput = HtmlService.createHtmlOutput('<p>Loading XML for MCQs...</p>')
    .setTitle('zyBooks AI Add-on')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(htmlOutput);

  var promptBase =  mcq_instructions_xml + mcq_guidelines + mcq_xml
  aiBlogBox(promptBase)
}

function testXML(){
  var htmlOutput = HtmlService.createHtmlOutput('<p>Loading XML for test bank questions...</p>')
    .setTitle('zyBooks AI Add-on')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(htmlOutput);

  var promptBase = test_instructions + test_xml
  aiBlogBox(promptBase)
}

function aiBlog(basePrompt) {
  var selectedText = doc.getSelection().getRangeElements()[0].getElement().asText().getText()
  // Replace YOUR_API_KEY with your actual OpenAI API key
  var apiKey = 'FIX_ME';
  var prompt = basePrompt + " " + selectedText;

  var model = "gpt-4-0125-preview"
  temperature = 1
  top_p = 0
  maxTokens = 4096
  seed = 123

    // Set up the request body with the given parameters
    const requestBody = {
      "model": model,
      "temperature": temperature,
      "max_tokens": maxTokens,
      "messages": [{"role": "user", "content": prompt}]
    };
    const requestOptions = {
      "method": "POST",
      "headers": {
        "Content-Type": "application/json",
        "Authorization": "Bearer "+ apiKey
      },
      "payload": JSON.stringify(requestBody)
    }

  // Call the OpenAI API
  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);

  // Parse the response and get the generated text
  var responseText = response.getContentText();
  var json = JSON.parse(responseText);
  Logger.log(json['choices'][0]['message']['content'])
  var output = json['choices'][0]['message']['content']
  var htmlOutput = HtmlService.createHtmlOutput('<p>Here are some suggestions that can be used as a starting point. Do not use these verbatim.</p><br/>' +  output)
    .setTitle('zyBooks AI Add-on')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(htmlOutput);
}

function aiBlogBox(basePrompt) {
  var selectedText = doc.getSelection().getRangeElements()[0].getElement().asText().getText()
  // Replace YOUR_API_KEY with your actual OpenAI API key
  var apiKey = 'FIX_ME';
  var prompt = basePrompt +" "+ selectedText;

  var model = "gpt-4-0125-preview"
  temperature = 1
  top_p = 0
  maxTokens = 4096
  seed = 123

    // Set up the request body with the given parameters
    const requestBody = {
      "model": model,
      "temperature": temperature,
      "max_tokens": maxTokens,
      "messages": [{"role": "user", "content": prompt}]
    };
    const requestOptions = {
      "method": "POST",
      "headers": {
        "Content-Type": "application/json",
        "Authorization": "Bearer "+ apiKey
      },
      "payload": JSON.stringify(requestBody)
    }

  // Call the OpenAI API
  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);

  // Parse the response and get the generated text
  var responseText = response.getContentText();
  var json = JSON.parse(responseText);
  Logger.log(json['choices'][0]['message']['content'])
  var output = json['choices'][0]['message']['content']
  var htmlOutput = HtmlService.createHtmlOutput('<p>Here are some suggestions that can be used as a starting point. Do not use these verbatim.</p><br/>' +  '<textarea rows="100" cols="35">' + output + '</textarea>')
    .setTitle('zyBooks AI Add-on')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(htmlOutput);
}
