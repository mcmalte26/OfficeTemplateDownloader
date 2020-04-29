using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Mime;
using System.Text.RegularExpressions;
using System.Threading;

namespace OfficeTemplateDownloader {
    public class Program {
        static void Main(string[] args) {
            Console.Title = "MS Office PowerPoint template downloader by MC_Malte";
            Dictionary<string, string> templates = new Dictionary<string, string>();

            Console.WriteLine("Instructions:");
            Console.WriteLine("Choose a template from https://templates.office.com/en-us/themes and copy paste it's link here \ne.g. https://templates.office.com/en-us/faded-pastoral-tm11531919");
            Console.WriteLine();

            while (true) {
                Console.WriteLine("Enter link or press ENTER to start downloading the templates");
                string input = Console.ReadLine();

                if (ContainsId(input)) {
                    string id = GetId(input);
                    Console.Write($"ID is {id}. Enter name for template: ");

                    string name = Console.ReadLine();
                    templates.Add(name ?? id, id);

                    Console.WriteLine();
                } else if (input == "") {
                    if (templates.Count == 0) {
                        Console.WriteLine("Error. You have to add at least one template to download.");
                        Console.WriteLine();
                    } else {
                        break;
                    }
                } else {
                    Console.WriteLine("Please enter a valid link.");
                }
            }

            Console.WriteLine();
            Console.WriteLine($"Added {templates.Count} templates to download. Press any key to start download . . .");
            Console.ReadLine();

            using WebClient client = new WebClient();
            foreach (KeyValuePair<string, string> template in templates) {
                string filename = $"{template.Key}.potx";

                Uri uri = new Uri(
                    $"https://omextemplates.content.office.net/support/templates/en-us/tf{template.Value}.potx");
                client.DownloadFileTaskAsync(uri, filename).Wait();

                Console.WriteLine($"Downloaded template {template.Key} with ID {template.Value} to file {filename}");
            }
        }

        private static readonly Regex IdRegex = new Regex(@"\d+");
        static bool ContainsId(string url) => IdRegex.Match(url).Success;
        static string GetId(string url) => IdRegex.Match(url).Value;
    }
}
