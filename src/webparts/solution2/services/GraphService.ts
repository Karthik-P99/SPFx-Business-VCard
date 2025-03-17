import { graphfi } from '@pnp/graph';
import { GraphFI } from '@pnp/graph/fi';
import "@pnp/graph/users";
import "@pnp/graph/cloud-communications";
import { IUserProfile } from '../models/IUserProfile';
import { IPersonaProps } from '@fluentui/react';

export default class GraphService {

  public static async getUserProfile(userIdentifier: string, graph: GraphFI = graphfi()): Promise<IUserProfile> {
    try {
      // Check if the identifier looks like an email address
      const isEmail = userIdentifier.indexOf('@') > -1;
      const userproperties = [
        'displayName', 'givenName', 'surname', 'mail', 'userPrincipalName', 'jobTitle', 
        'department', 'officeLocation', 'mobilePhone', 'businessPhones', 
        'preferredLanguage', 'id', 'companyName', 'accountEnabled'
      ];

      // Fetch user data
      let user;
      if (isEmail) {
        const users = await graph.users
          .filter(`mail eq '${userIdentifier}' or userPrincipalName eq '${userIdentifier}'`)
          .select(userproperties.join(","))
          .top(1)();
        
        if (!users || users.length === 0) {
          throw new Error(`User with email ${userIdentifier} not found`);
        }
        user = users[0];
      } else {
        user = await graph.users.getById(userIdentifier).select(userproperties.join(","))();
      }

      // Try to get user photo
      let photoUrl = '';
      try {
        if (user.id) {
          const photo = await graph.users.getById(user.id).photo.getBlob();
          if (photo) photoUrl = URL.createObjectURL(photo);
        }
      } catch (photoError) {
        console.log('User photo not available:', photoError);
      }

      // Try to get presence
      let presence = null;
      try {
        if (user.id) presence = await graph.users.getById(user.id).presence();
      } catch (presenceError) {
        console.log('User presence not available:', presenceError);
      }

      // Build and return profile
      return {
        displayName: user.displayName || '',
        givenName: user.givenName || '',
        surname: user.surname || '',
        mail: user.mail || user.userPrincipalName || '',
        userPrincipalName: user.userPrincipalName || '',
        jobTitle: user.jobTitle || '',
        department: user.department || '',
        officeLocation: user.officeLocation || '',
        mobilePhone: user.mobilePhone || '',
        businessPhones: user.businessPhones || [],
        preferredLanguage: user.preferredLanguage || '',
        photo: photoUrl,
        id: user.id || '',
        companyName: user.companyName || '',
        accountEnabled: user.accountEnabled || false,
        presence: presence && presence.activity === 'OffWork' ? 'OffWork' : 
             (presence && presence.availability ? presence.availability : 'Unknown')
      };
    } catch (error) {
      console.error('Error in getUserProfile:', error);
      throw error;
    }
  }

  /**
   * Search users based on a filter term
   */
  public static searchUsers(filter: string, client: any): Promise<IPersonaProps[]> {
    if (!filter || filter.length < 2) return Promise.resolve([]);

    return client
      .api('/users')
      .filter(`startswith(displayName,'${filter}') or startswith(userPrincipalName,'${filter}')`)
      .select('id,displayName,userPrincipalName,mail,jobTitle')
      .top(10)
      .get()
      .then((response: any) => response.value.map((user: any) => ({
        text: user.displayName,
        secondaryText: user.mail || user.userPrincipalName,
        tertiaryText: user.jobTitle,
        id: user.id
      })))
      .catch((error: any) => {
        console.error('Error searching for people:', error);
        throw error;
      });
  }

  /**
   * Get multiple user profiles at once
   */
  public static async getUserProfiles(userIds: string[], graph: GraphFI = graphfi()): Promise<IUserProfile[]> {
    if (!userIds || userIds.length === 0) return [];
    
    try {
      return await Promise.all(userIds.map(userId => this.getUserProfile(userId, graph)));
    } catch (error) {
      console.error('Error in getUserProfiles:', error);
      throw error;
    }
  }
}