using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Sitecore.Mvc;
using Sitecore.Mvc.Presentation;
using Sitecore.Data;
using Sitecore.Data.Items;

namespace GameStoreProject.Feature.PageContent.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            Database database = Sitecore.Context.Database;
            Item gamesFolder = database.GetItem("{F7916301-8BA7-41D5-A9B4-CEDB5B486249}");
            List<Item> games = gamesFolder.GetChildren().ToList();
            return View(games);
        }
    }
}