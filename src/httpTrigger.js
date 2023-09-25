const notificationTemplate = require("./adaptiveCards/notification-default.json");
const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
const { notificationApp } = require("./internal/initialize");
const { NotificationTargetType } = require("@microsoft/teamsfx");

// HTTP trigger to send notification. You need to add authentication / authorization for this API. Refer https://aka.ms/teamsfx-notification for more details.
module.exports = async function (context, req) {
  console.log();
  const notificationUserId = context.req.body.user_id;
  const notificationTitle = context.req.body.title;
  const notificationMessage = context.req.body.message;

  const pageSize = 100;
  let continuationToken = undefined;
  do {
    const pagedData = await notificationApp.notification.getPagedInstallations(
      pageSize,
      continuationToken
    );
    const installations = pagedData.data;
    continuationToken = pagedData.continuationToken;

    for (const target of installations) {

      if (target.type === NotificationTargetType.Person) {
        // Directly notify the individual person
        const member = await notificationApp.notification.findMember(
          async (m) => m.account.aadObjectId === notificationUserId
        );
        await member?.sendAdaptiveCard(
          AdaptiveCards.declare(notificationTemplate).render({
            title: notificationTitle,
            appName: "Speybl",
            description: notificationMessage,
            notificationUrl: "https://speybl.com",
          })
        );
      }
    }
  } while (continuationToken);
  context.res = {};
};
