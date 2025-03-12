using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using Markdig;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFunctions.AI
{
    public static class AIFunctions
    {
        [ExcelFunction(Category = "String", Description = "Sends a user chat text message to the model and returns the response messages.", HelpTopic = "Sends a user chat text message to the model and returns the response messages.")]
        public static object AICHAT(
[ExcelArgument("message", Name = "message", Description = "The chat content to send.")] string message)
        {
            try
            {
                string functionName = nameof(AI);
                object[] parameters = new object[] { message };

                return AsyncTaskUtil.RunTask<object>(functionName, parameters, async () =>
                {
                    //The actual asyncronous block of code to execute.
                    IChatClient _chatClient = ContainerOperations.Container.GetRequiredService<IChatClient>();
                    ChatMessage chatMessage = new ChatMessage(ChatRole.User, message);
                    ChatResponse response = await _chatClient.GetResponseAsync(message);

                    return response.Text;
                });
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }
    }
}
