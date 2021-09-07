import React, { useEffect, useState } from 'react';
import styles from './SideNav.module.scss';
import { Nav, INavLink } from 'office-ui-fabric-react/lib/Nav';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/navigation";
import { ISerializableNavigationNode } from '@pnp/sp/navigation';

const SideNav = () => {
    const [navItems, setNavItems] = useState<INavLink[]>([]);

    useEffect(() => {
        const asyncFn = async () => {
            const items = await getSideNavigationItems();
            setNavItems(items);
        };
        asyncFn();
    }, []);

    const getSideNavigationItems = async (): Promise<INavLink[]> => {
        const sideNavItems: INavLink[] = [];
 
        const quickLaunchItems: ISerializableNavigationNode[] = await sp.web.navigation.quicklaunch.expand("Children").get();

        quickLaunchItems.forEach((item: ISerializableNavigationNode): void => {
            sideNavItems.push({
                name: item.Title,
                url: item.Url,
                key: item.Id.toString(),
                links: getNavItems(item.Children)
            });
        });

        return sideNavItems;
    };

    const getNavItems = (items: ISerializableNavigationNode[]) => {
        const subNavItems: INavLink[] = [];
        items.forEach((item: ISerializableNavigationNode): void => {
            subNavItems.push({
                key: item.Id.toString(),
                name: item.Title,
                url: item.Url
            });
        });
        return subNavItems;
    };

    return (
        <div className={styles.spfxFluentuiNav} >
            <Nav
                ariaLabel="Navigation"
                groups={[{ links: navItems }]} />
        </div>
    );
};

export default SideNav;