Pour faire fonctionner un Office Add-in, il faut déployer le manifest sur le tenant cible et déployer le projet dans un azure storage static (en utilisant vs code, clic droit sur "dist" puis "deploy to static website via azure storage"). 

Pour une maj, pas besoin de changer le manifest côté client, il faut juste re déployer les changements dans l'azure storage (sauf dans le cas d'un changements de nom de l'appli côté client, il faut mettre à jour la version du et mettre à jour depuis le panneau d'administration office via le menu "applications intégrées")

Dans le cas d'un changement de storage azure (nouveau dépôt, ...), il faudra aussi mettre à jour le manifest chez le client en changeant les url faisant référence à l'azure storage (commentaires devant) et changer le numéro de version.

Pour avoir plus d'informations sur le projet, voir dans l'équipe Teams Traitement des eaux de Cabinet Merlin.
 