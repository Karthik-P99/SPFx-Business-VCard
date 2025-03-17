import * as React from 'react';
import { ISolution2Props } from './ISolution2Props';
import BusinessCard from './businesscard/BusinessCard';
import styles from './Solution2.module.scss';
import { Pivot, PivotItem } from '@fluentui/react';

const Solution2: React.FC<ISolution2Props> = (props) => {
  return (
    <div className={styles.solution2}>
      <Pivot>
        <PivotItem headerText="Business Card" itemKey="businessCard">
          <div className={styles.content}>
            <BusinessCard context={props.context} description={props.description} />
          </div>
        </PivotItem>
      </Pivot>
    </div>
  );
};

export default Solution2;
