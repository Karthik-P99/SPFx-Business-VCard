import * as React from 'react';
import { useState, useEffect } from 'react';
import { graphfi } from '@pnp/graph';
import '@pnp/graph/users';
import '@pnp/graph/photos';
import { SPFx } from '@pnp/graph';
import { IUserProfile } from '../../models/IUserProfile';
import GraphService from '../../services/GraphService';
import { QRCodeSVG } from 'qrcode.react';
import styles from './UserCard.module.scss';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import {
  Persona,
  PersonaSize,
  PersonaPresence,
  IPersonaProps,
  IPersonaStyles,
  IBasePickerSuggestionsProps,
  ListPeoplePicker,
  ValidationState,
  PeoplePickerItemSuggestion
} from '@fluentui/react';
import { Icon } from '@fluentui/react/lib/Icon';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// UserCard props interface
interface IUserCardProps {
  context: WebPartContext; 
  userId?: string;
}

// Custom persona styles for wrapped text
const personaStyles: Partial<IPersonaStyles> = {
  root: { height: 'auto' },
  secondaryText: { height: 'auto', whiteSpace: 'normal' },
  primaryText: { height: 'auto', whiteSpace: 'normal' },
};

// Helper functions for people picker
const getTextFromItem = (persona: IPersonaProps): string => persona.text as string;
const validateInput = (input: string): ValidationState =>
  input.indexOf('@') !== -1 ? ValidationState.valid :
    input.length > 1 ? ValidationState.warning : ValidationState.invalid;

// Mapping function to convert Graph API presence values to PersonaPresence values
const mapPresenceToPersonaPresence = (presence: string): PersonaPresence => {
  switch (presence) {
    case 'Available':
      return PersonaPresence.online;
    case 'Busy':
      return PersonaPresence.busy;
    case 'DoNotDisturb':
      return PersonaPresence.dnd;
    case 'Away':
      return PersonaPresence.away;
    case 'Offline':
      return PersonaPresence.offline;
    case 'BeRightBack':
      return PersonaPresence.away;
    case 'InACall':
    case 'InAConferenceCall':
      return PersonaPresence.busy;
    case 'Presenting':
      return PersonaPresence.busy;
    case 'OutOfOffice':
      return PersonaPresence.offline;
    default:
      return PersonaPresence.none;
  }
};

// UserCard as functional component
const UserCard: React.FC<IUserCardProps> = ({ context, userId: initialUserId }) => {
  const [userProfiles, setUserProfiles] = useState<IUserProfile[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [selectedUserIds, setSelectedUserIds] = useState<string[]>(initialUserId ? [initialUserId] : []);
  const [hasSearched, setHasSearched] = useState<boolean>(false);
  const [selectedPeople, setSelectedPeople] = useState<IPersonaProps[]>([]);

  // Generate vCard format for QR code
  const generateVCardData = (profile: IUserProfile): string => {
    if (!profile) return '';

    const vCardFields = [
      'BEGIN:VCARD',
      'VERSION:3.0',
      `FN:${profile.displayName || ''}`,
      `N:${profile.surname || ''};${profile.givenName || ''};;;`,
      `TITLE:${profile.jobTitle || ''}`,
      `ORG:${profile.companyName || ''};${profile.department || ''}`,
      `EMAIL:${profile.mail || ''}`,
      `TEL;TYPE=WORK,VOICE:${profile.businessPhones && profile.businessPhones.length > 0 ? profile.businessPhones[0] : ''}`,
      `TEL;TYPE=CELL,VOICE:${profile.mobilePhone || ''}`,
      `ADR;TYPE=WORK:;;${profile.officeLocation || ''};;;;`,
      'END:VCARD'
    ];

    return vCardFields.join('\n');
  };

  useEffect(() => {
    const fetchUserData = async (): Promise<void> => {
      try {
        setLoading(true);
        setError(null);

        const graph = graphfi().using(SPFx(context));
        const profiles = await GraphService.getUserProfiles(selectedUserIds, graph);

        setUserProfiles(profiles);
        setHasSearched(true);
      } catch (e) {
        console.error("Error fetching user profiles:", e);
        setError('Failed to load user profiles');
      } finally {
        setLoading(false);
      }
    };

    // Only fetch if we have selectedUserIds
    if (selectedUserIds.length > 0) {
      // Use void operator to explicitly mark that we're ignoring the promise
      void fetchUserData();
    } else if (initialUserId) {
      setSelectedUserIds([initialUserId]);
    } else {
      setHasSearched(false);
    }
  }, [context, selectedUserIds, initialUserId]);

  // Handler for when people picker selection changes
  const onPeoplePickerChange = (items?: IPersonaProps[]): void => {
    setSelectedPeople(items || []);

    if (items && items.length > 0) {
      // Extract user IDs from email addresses (preferably) or display names
      const userIds = items.map(item => item.secondaryText || item.text || '');
      setSelectedUserIds(userIds);
      setHasSearched(true);
    } else {
      // Reset state if no selection
      setSelectedUserIds([]);
      setUserProfiles([]);
      setHasSearched(false);
    }
  };

  // People picker suggestions props
  const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested People',
    mostRecentlyUsedHeaderText: 'Suggested Contacts',
    noResultsFoundText: 'No results found',
    loadingText: 'Loading...',
    showRemoveButtons: true,
    suggestionsAvailableAlertText: 'People Picker Suggestions available',
    suggestionsContainerAriaLabel: 'Suggested contacts',
  };

  // Function to resolve input to personas for people picker
  const onResolveSuggestions = (filter: string): Promise<IPersonaProps[]> => {
    if (!filter || filter.length < 2) return Promise.resolve([]);

    return new Promise<IPersonaProps[]>((resolve, reject) => {
 
      context.msGraphClientFactory
        .getClient("3")
        .then(
          (client: MSGraphClientV3) => 
            GraphService.searchUsers(filter, client)
              .then(resolve)
              .catch(reject),
    
          (error: Error) => {
            console.error("Error getting Graph client:", error);
            reject(error);
          }
        );
    });
  };

  // Custom renderer for suggestion items with wrapped text
  const onRenderSuggestionItem = (personaProps: IPersonaProps, suggestionsProps: IBasePickerSuggestionsProps): JSX.Element =>
    <PeoplePickerItemSuggestion
      personaProps={{ ...personaProps, styles: personaStyles }}
      suggestionsProps={suggestionsProps}
    />;

  return (
    <div className={styles.container}>
      <div className={styles.peoplePickerContainer}>
        <ListPeoplePicker
          onResolveSuggestions={onResolveSuggestions}
          onEmptyInputFocus={() => Promise.resolve([])}
          getTextFromItem={getTextFromItem}
          pickerSuggestionsProps={suggestionProps}
          className={'ms-PeoplePicker'}
          onChange={onPeoplePickerChange}
          selectedItems={selectedPeople}
          key={'list'}
          selectionAriaLabel={'Selected users'}
          removeButtonAriaLabel={'Remove'}
          onRenderSuggestionsItem={onRenderSuggestionItem}
          onValidateInput={validateInput}
          inputProps={{
            placeholder: 'Enter names or emails',
            'aria-label': 'Search Users'
          }}
          resolveDelay={300}
          itemLimit={5}
        />
      </div>

      {loading ? (
        <div className={styles.loadingContainer}>Loading user information...</div>
      ) : hasSearched && error ? (
        <div className={styles.errorContainer}>Error: {error}</div>
      ) : userProfiles.length > 0 ? (
        <div className={styles.cardsContainer}>
          {userProfiles.map((profile, index) => (
            <div key={index} className={styles.businessCard}>
              <div className={styles.cardContent}>
                <div className={styles.leftSection}>
                  <div className={styles.photoContainer} title={profile.presence ? `Status: ${profile.presence}` : 'Status: Offline'}>
                    <Persona
                      imageUrl={profile.photo}
                      imageAlt={`${profile.displayName}`}
                      text={profile.displayName}
                      size={PersonaSize.size72}
                      hidePersonaDetails={true}
                      presence={profile.presence ? mapPresenceToPersonaPresence(profile.presence) : PersonaPresence.offline}
                      showInitialsUntilImageLoads={true}
                      imageInitials={profile.displayName?.split(' ')
                        .map(name => name.charAt(0).toUpperCase())
                        .join('')
                        .substring(0, 2)}
                    />
                  </div>
                  <div className={styles.userInfo}>
                    <h2>{profile.displayName}</h2>
                    {profile.jobTitle && <p className={styles.jobTitle}>{profile.jobTitle}</p>}

                    <div className={styles.detailsGrid}>
                      {profile.accountEnabled !== undefined && (
                        <div className={styles.gridItem}>
                          
                          <span><strong>Account Enabled:</strong> </span>
                          <Icon
                            iconName={profile.accountEnabled ? "CheckMark" : "Cancel"}
                            className={profile.accountEnabled ? styles.enabledIcon : styles.disabledIcon}
                          />
                        </div>
                      )}

                      {profile.companyName && (
                        <div className={styles.gridItem}>
                          <Icon iconName="Bank" className={styles.detailIcon} />
                          <span><strong>Company:</strong> {profile.companyName}</span>
                        </div>
                      )}

                      {profile.department && (
                        <div className={styles.gridItem}>
                          <Icon iconName="Org" className={styles.detailIcon} />
                          <span><strong>Department:</strong> {profile.department}</span>
                        </div>
                      )}

                      {profile.officeLocation && (
                        <div className={styles.gridItem}>
                          <Icon iconName="Location" className={styles.detailIcon} />
                          <span><strong>Office:</strong> {profile.officeLocation}</span>
                        </div>
                      )}

                      {profile.mail && (
                        <div className={styles.gridItem}>
                          <Icon iconName="Mail" className={styles.detailIcon} />
                          <span><strong>Email:</strong> {profile.mail}</span>
                        </div>
                      )}

                      {profile.mobilePhone && (
                        <div className={styles.gridItem}>
                          <Icon iconName="CellPhone" className={styles.detailIcon} />
                          <span><strong>Mobile:</strong> {profile.mobilePhone}</span>
                        </div>
                      )}

                      {profile.businessPhones && profile.businessPhones.length > 0 && (
                        <div className={styles.gridItem}>
                          <Icon iconName="Phone" className={styles.detailIcon} />
                          <span><strong>Business Phone:</strong> {profile.businessPhones[0]}</span>
                        </div>
                      )}
                    </div>

                  </div>
                </div>
                <div className={styles.qrCodeSection}>
                  <QRCodeSVG
                    value={generateVCardData(profile)}
                    size={180}
                    includeMargin={true}
                  />
                  <p className={styles.qrCodeLabel}>Scan to add contact</p>
                </div>
              </div>
            </div>
          ))}
        </div>
      ) : hasSearched ? (
        <div className={styles.infoContainer}>No user information available</div>
      ) : (
        <div className={styles.infoContainer}>Search for users to display their information</div>
      )}
    </div>
  );
};

export default UserCard;