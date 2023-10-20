const { TeamsActivityHandler } = require("botbuilder");

// An empty teams activity handler.
// You can add your customization code here to extend your bot logic if needed.
class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
    this.onMembersAdded(async (context, next) => {
      for (const idx in context.activity.membersAdded) {
        if (context.activity.membersAdded[idx].id !== context.activity.recipient.id) {
          await context.sendActivity(`👋 Welcome to Speybl Notification Bot for Microsoft Teams!`)
          await context.sendActivity(`Speybl is here to keep you informed and updated with important notifications, right within your Microsoft Teams workspace. Please note that this bot is designed to provide one-way notifications, ensuring you never miss out on critical information.`);
          await context.sendActivity(`To get started, register on our application at [app.speybl.com](https://app.speybl.com)`);
          await context.sendActivity(`Feel free to reach out to us if you have any questions, concerns, or need assistance. Our dedicated support team is here to help you. You can contact us at [hello@speybl.com](mailto:hello@speybl.com).`);
        }
      }
      await next();
    });
  }
}

module.exports.TeamsBot = TeamsBot;
