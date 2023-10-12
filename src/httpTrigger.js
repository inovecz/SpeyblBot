const simpleTemplate = require("./adaptiveCards/notification-simple.json");
const linkTemplate = require("./adaptiveCards/notification-link.json");
const twoActionsTemplate = require("./adaptiveCards/notification-two-actions.json");
const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
const { notificationApp } = require("./internal/initialize");
const { NotificationTargetType } = require("@microsoft/teamsfx");

// HTTP trigger to send notification. You need to add authentication / authorization for this API. Refer https://aka.ms/teamsfx-notification for more details.
module.exports = async function (context, req) {
  const notificationType = context.req.body.type ?? 'simple'
  const notificationUserId = context.req.body.user_id;
  const notificationTitle = context.req.body.title;
  const notificationMessage = context.req.body.message;
  const notificationActions = context.req.body.actions

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
        
          let notificationData = {
            title: notificationTitle,
            appName: "Speybl",
            description: notificationMessage,
            actions: notificationActions
          }
          
          if(notificationType === 'simple') {
            await member?.sendAdaptiveCard(
              AdaptiveCards.declare(simpleTemplate).render(notificationData)
            );
          } else if (notificationType === 'link') {
            if(notificationActions && notificationActions.length >= 1) {
              await member?.sendAdaptiveCard(
                AdaptiveCards.declare(linkTemplate).render(notificationData)
              );
            }
            
          }  else if (notificationType === 'two-actions') {
            if(notificationActions && notificationActions.length >= 2) {
              console.log('here');
              await member?.sendAdaptiveCard(
                AdaptiveCards.declare(twoActionsTemplate).render(notificationData)
              );
            }
          }
      }
    }
  } while (continuationToken);
  context.res = {};
};
