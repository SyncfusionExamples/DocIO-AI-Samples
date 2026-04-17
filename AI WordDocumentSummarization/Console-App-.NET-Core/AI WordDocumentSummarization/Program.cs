using OpenAI;
using OpenAI.Chat;
using Syncfusion.DocIO.DLS;
using System.ClientModel;

namespace AI_WordDocumentSummarization
{
    class WordDocumentSummarizer
    {
        // Replace with your actual OpenAI API key or set it in environment variables for security
        static string? openAIApiKey = "Replace the OpenAI Key";
        static string? openAImodel = "gpt-4o-mini";
        static async Task Main()
        {
            await ExecuteSummarization();
        }
        /// <summary>
        /// Execute summarization of Word document.
        /// </summary>
        private async static Task ExecuteSummarization()
        {
            Console.WriteLine("AI Powered Word Summarizer");

            Console.WriteLine("Enter full Word file path (e.g., C:\\Data\\Input.docx):");

            //Read user input for Word file path
            string? wordFilePath = Console.ReadLine()?.Trim().Trim('"');

            Console.WriteLine("Please enter the number of sentences you would like the summary to be (e.g., 3, 5):");

            //Read user input for required number of lines
            string? sentencesCount = Console.ReadLine()?.Trim().Trim('"');

            if (string.IsNullOrWhiteSpace(wordFilePath) || !File.Exists(wordFilePath))
            {
                Console.WriteLine("Invalid path. Exiting.");
                return;
            }

            if (string.IsNullOrWhiteSpace(sentencesCount) || !int.TryParse(sentencesCount, out int result))
            {
                Console.WriteLine("Invalid Count. Exiting.");
                return;
            }

            if (string.IsNullOrWhiteSpace(openAIApiKey))
            {
                Console.WriteLine("OPENAI_API_KEY not set. Exiting.");
                return;
            }

            try
            {
                // Summarize Word content
                await SummarizeWordContent(wordFilePath, sentencesCount);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to summarize Word document: {ex.StackTrace}");
                return;
            }
        }

        private static async Task SummarizeWordContent(string wordFilePath, string sentencesCount)
        {
            // Open the Existing Word document
            WordDocument wordDocument = new WordDocument(wordFilePath);

            // Define the system prompt that instructs OpenAI how to summarize the content,
            // including the desired number of sentences in the summary.
            string systemPrompt = @"You are a professional document summarizer integrated into an DocIO automation tool.
                                    Your job is to summarize the word document content into the" + sentencesCount + " sentences";
            string originalText = wordDocument.GetText();
            wordDocument.Close();
            // Call OpenAI to summarize the text and store the summarized result
            string summarizedText = await AskOpenAIAsync(openAIApiKey, openAImodel, systemPrompt, originalText);

            WordDocument summarizedDocument = new WordDocument();
            summarizedDocument.EnsureMinimal();
            summarizedDocument.LastParagraph.AppendText(summarizedText);
            summarizedDocument.Save(wordFilePath.Replace(".docx", "_DocIOsummarized.docx"));
            summarizedDocument.Close();
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