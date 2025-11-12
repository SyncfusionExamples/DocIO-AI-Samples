using OpenAI;
using OpenAI.Chat;
using Syncfusion.DocIO.DLS;
using System.ClientModel;

namespace AI_WordDocumentTranslator
{
    class WordDocumentTranslator
    {
        // Replace with your actual OpenAI API key or set it in environment variables for security
        static string? openAIApiKey = "Replace the OpenAI Key";
        static string? openAImodel = "gpt-4o-mini";
        static async Task Main()
        {
            await ExecuteTranslation();
        }
        /// <summary>
        /// Execute translation of Word document.
        /// </summary>
        private async static Task ExecuteTranslation()
        {
            Console.WriteLine("AI Powered Word Translator");

            Console.WriteLine("Enter full Word file path (e.g., C:\\Data\\Input.docx):");

            //Read user input for Word file path
            string? wordFilePath = Console.ReadLine()?.Trim().Trim('"');

            Console.WriteLine("Enter the language name (e.g., Chinese, Japanese)");

            //Read user input for required language
            string? language = Console.ReadLine()?.Trim().Trim('"');

            if (string.IsNullOrWhiteSpace(wordFilePath) || !File.Exists(wordFilePath))
            {
                Console.WriteLine("Invalid path. Exiting.");
                return;
            }

            if (string.IsNullOrWhiteSpace(language))
            {
                Console.WriteLine("Invalid language. Exiting.");
                return;
            }

            if (string.IsNullOrWhiteSpace(openAIApiKey))
            {
                Console.WriteLine("OPENAI_API_KEY not set. Exiting.");
                return;
            }

            try
            {
                // Translate Word content
                await TranslateWordContent(wordFilePath, language);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read Word document: {ex.StackTrace}");
                return;
            }
        }
        /// <summary>
        /// Translate Word content using OpenAI and Syncfusion DocIO.
        /// </summary>
        /// <param name="openAIApiKey">OpenAI API key.</param>
        /// <param name="wordFilePath">Path to the Word file.</param>
        /// <param name="language">Target language for translation.</param>
        private static async Task TranslateWordContent(string wordFilePath, string language)
        {
            // Open the Existing Word document
            WordDocument wordDocument = new WordDocument(wordFilePath);

            // Define the system prompt that instructs OpenAI how to translate
            // Includes rules to preserve formatting, placeholders, and avoid paraphrasing
            string systemPrompt = @"You are a professional translator integrated into an DocIO automation tool.
                                    Your job is to translate text from word document into the" + language + @" language
                                    Rules:
                                    - Preserve original structure as much as possible.
                                    - Prefer literal translation over paraphrasing.
                                    - Return ONLY the translated text, without quotes, labels, or explanations.
                                    - Preserve placeholders (e.g., {0}, {name}) and keep numbers, currency, and dates intact.
                                    - Do not change the meaning, tone, or formatting unnecessarily.
                                    - Do not add extra commentary or code fences.
                                    - If the text is already in the target language, return it unchanged.
                                    - Be concise and accurate.";

            // Loop through each section in the Word document
            foreach (WSection section in wordDocument.Sections)
            {
                //Accesses the Body of section where all the contents in document are apart
                WTextBody sectionBody = section.Body;
                await TranslateTextBody(sectionBody, systemPrompt);
                WHeadersFooters headersFooters = section.HeadersFooters;
                //Consider that OddHeader and OddFooter are applied to this document
                //Iterates through the TextBody of OddHeader and OddFooter
                await TranslateTextBody(headersFooters.OddHeader, systemPrompt);
                await TranslateTextBody(headersFooters.OddFooter, systemPrompt);
            }
            // Save the Translate Word document.
            wordDocument.Save(wordFilePath.Replace(".docx", "_DocIOTranslate.docx"));
            wordDocument.Close();
        }
        /// <summary>
        /// Translates the text content of a text body using OpenAI and updates it.
        /// </summary>
        /// <param name="textBody">The textBody to translate.</param>
        /// <param name="systemPrompt">System prompt to guide translation behavior.</param>
        private static async Task TranslateTextBody(WTextBody textBody, string systemPrompt)
        {
            // Loop through each entity (paragraph, table, etc.) in the section body.
            foreach (Entity entity in textBody.ChildEntities)
            {
                // If the entity is a paragraph, call the paragraph translation method
                if (entity is WParagraph paragraph)
                {
                    await TranslateParagraphs(paragraph, systemPrompt);
                }
                // If the entity is a table, call the table translation method
                else if (entity is WTable table)
                {
                    await TranslateTable(table, systemPrompt);
                }
                else if (entity is BlockContentControl blockContentControl)
                {
                    await TranslateTextBody(blockContentControl.TextBody, systemPrompt);
                }
            }
        }
        /// <summary>
        /// Translates the text content of a paragraph using OpenAI and updates it.
        /// </summary>
        /// <param name="paragraph">The paragraph to translate.</param>
        /// <param name="systemPrompt">System prompt to guide translation behavior.</param>
        private static async Task TranslateParagraphs(WParagraph paragraph, string systemPrompt)
        {
            string originalText = paragraph.Text;
            if (string.IsNullOrEmpty(originalText)) return;
            try
            {
                string translatedText = originalText;
                // Call OpenAI to translate the text and Store the translated result in the dictionary for reuse
                translatedText = await AskOpenAIAsync(openAIApiKey, openAImodel, systemPrompt, originalText);
                // Replace the original text with the translated version
                paragraph.Text = translatedText;
            }
            catch (Exception ex)
            {
                // Log any errors that occur during translation
                Console.WriteLine($"OpenAI error: {ex.Message}");
            }
        }
        /// <summary>
        /// Translates all text content within a table using OpenAI.
        /// </summary>
        /// <param name="table">The table to translate.</param>
        /// <param name="systemPrompt">System prompt to guide translation behavior.</param>
        private static async Task TranslateTable(WTable table, string systemPrompt)
        {
            // Loop through each row in the table
            foreach (WTableRow row in table.Rows)
            {
                // Loop through each cell in the current row
                foreach (WTableCell cell in row.Cells)
                {
                    // Loop through each entity (paragraph or nested table) inside the cell
                    foreach (Entity entity in cell.ChildEntities)
                    {
                        // If the entity is a paragraph, call the paragraph translation method
                        if (entity is WParagraph paragraph)
                        {
                            await TranslateParagraphs(paragraph, systemPrompt);
                        }
                        // If the entity is a nested table, call the table translation method
                        else if (entity is WTable nestedTable)
                        {
                            await TranslateTable(nestedTable, systemPrompt);
                        }
                       
                    }
                }
            }
        }
        /// <summary>
        /// Sends a chat completion request to OpenAI and returns the response.
        /// </summary>
        /// <param name="apiKey">OpenAI API key.</param>
        /// <param name="model">Model name.</param>
        /// <param name="systemPrompt">System prompt.</param>
        /// <param name="userContent">User content.</param>
        /// <returns>AI-generated response as a string.</returns>
        private static async Task<string> AskOpenAIAsync(string apiKey, string model, string systemPrompt, string userContent)
        {
            // Initialize OpenAI client
            OpenAIClient openAIClient = new OpenAIClient(apiKey);

            // Create chat client for the specified model
            ChatClient chatClient = openAIClient.GetChatClient(model);

            //Get AI response
            ClientResult<ChatCompletion> chatResult = await chatClient.CompleteChatAsync(new SystemChatMessage(systemPrompt), new UserChatMessage(userContent));

            string response = chatResult.Value.Content[0].Text ?? string.Empty;

            return response;
        }
    }
}
