export interface Delivery {
  id: string;
  receiptNumber: string;
  company: string;
  customer: string;
  destination: string;
  locality: string;
  promisedDate: Date;
  creationDate: Date;
  status: 'Pendiente' | 'En Camino' | 'Demorado' | 'Entregado';
  carrier?: string;
  priority: 'Baja' | 'Media' | 'Alta' | 'Crítica';
}

export interface DeliveryStats {
  totalPending: number;
  totalDelayed: number;
  averageDelayDays: number;
  criticalDeliveries: number;
}
