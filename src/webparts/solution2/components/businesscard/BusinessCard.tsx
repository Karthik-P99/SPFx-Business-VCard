import * as React from 'react';
import { IBusinessCardProps } from './BusinessCardProps';
import UserCard from './UserCard';
import styles from './CardStyles.module.scss';

const BusinessCard: React.FC<IBusinessCardProps> = (props) => {
  return (
    <div className={styles.businessCard}>
      <UserCard context={props.context} />
    </div>
  );
};

export default BusinessCard;
