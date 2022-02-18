using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Automation.Ui.Accelerators.ReportingClassess
{
    public class Chapter
    {
        /// <summary>
        /// Creates a new Chapter
        /// </summary>
        /// <param name="title">Title</param>
        public Chapter(string title)
        {
            Title = title;
            Console.WriteLine($"***** {Title} *****");
        }

        /// <summary>
        /// Gets or sets Title
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets Steps
        /// </summary>
        public List<Step> Steps { get; } = new List<Step>();
        
        /// <summary>
        /// Get current Step
        /// </summary>
        public Step Step
        {
            get
            {
                if (Steps.Count() == 0)
                    Steps.Add(new Step("UNKNOWN STEP"));
                return Steps.Last();
            }
        }

        /// <summary>
        /// Gets or sets IsSuccess
        /// </summary>
        public bool IsSuccess
        {
            get
            {
                return Steps.Where(x => x.IsSuccess == false).Count() == 0 ? true : false;
            }

        }
    }
}