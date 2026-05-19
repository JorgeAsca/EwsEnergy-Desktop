import * as React from "react";
import { Nav, INavLinkGroup, INavLink, Text } from "@fluentui/react";
import styles from "../Obras.module.scss";

interface ISidebarProps {
  selectedKey: string;
  onLinkClick: (key: string) => void;
  isOpen?: boolean;
}

const navGroups: INavLinkGroup[] = [
  {
    links: [
      { name: "Inventario", url: "", key: "inventario", icon: "Package" },
      { name: "Personal", url: "", key: "personal", icon: "Group" },
      { name: "Obras", url: "", key: "obras", icon: "ConstructionCone" },
      { name: "Planificación", url: "", key: "planificacion", icon: "Calendar" }, 
      { name: "Asignaciones", url: "", key: "asignaciones", icon: "ContactLink" },
      { name: "Diario", url: "", key: "fotos", icon: "Camera" },
      { name: "Control de Obras", url: "", key: "historial", icon: "History" },
    ],
  },
];

export const Sidebar: React.FC<ISidebarProps> = (props) => {
  return (
    <div className={`${styles.sidebar} ${props.isOpen ? styles.isOpen : ""}`}>
      <div className={styles.logoArea}>
        <Text variant="large" style={{ fontWeight: "bold", color: "white" }}>
          EWS ENERGY
        </Text>
      </div>
      <Nav
        selectedKey={props.selectedKey}
        groups={navGroups}
        onLinkClick={(ev, item?: INavLink) => {
          if (ev) ev.preventDefault();
          if (item) {
            props.onLinkClick(item.key as string);
          }
        }}
      />
    </div>
  );
};