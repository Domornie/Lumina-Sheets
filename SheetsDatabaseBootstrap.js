function initializeSheetsDatabase() {
  if (typeof SheetsDB === 'undefined') {
    console.warn('SheetsDB is not available; skipping database bootstrap');
    return;
  }

  try {
    SheetsDB.defineTable({
      name: 'Customers',
      version: 1,
      primaryKey: 'id',
      idPrefix: 'CUS_',
      columns: [
        { name: 'id', type: 'string', primaryKey: true },
        { name: 'tenantId', type: 'string', required: true, index: true },
        { name: 'email', type: 'string', required: true, unique: true, index: true, pattern: '^[^@\s]+@[^@\s]+\\.[^@\s]+$' },
        { name: 'firstName', type: 'string', required: true, minLength: 1, maxLength: 100 },
        { name: 'lastName', type: 'string', required: true, minLength: 1, maxLength: 100 },
        { name: 'phone', type: 'string', nullable: true, maxLength: 32 },
        { name: 'status', type: 'enum', required: true, allowedValues: ['prospect', 'active', 'inactive'], defaultValue: 'prospect' },
        { name: 'notes', type: 'json', nullable: true }
      ],
      indexes: [
        { name: 'Customers_email', field: 'email' },
        { name: 'Customers_tenant', field: 'tenantId' }
      ],
      retentionDays: 365
    });

    SheetsDB.defineTable({
      name: 'Orders',
      version: 1,
      primaryKey: 'id',
      idPrefix: 'ORD_',
      columns: [
        { name: 'id', type: 'string', primaryKey: true },
        { name: 'tenantId', type: 'string', required: true, index: true },
        { name: 'customerId', type: 'string', required: true, references: { table: 'Customers', column: 'id', allowNull: false } },
        { name: 'orderTotal', type: 'number', required: true, min: 0 },
        { name: 'currency', type: 'enum', required: true, allowedValues: ['USD', 'EUR', 'GBP', 'PHP'], defaultValue: 'USD' },
        { name: 'status', type: 'enum', required: true, allowedValues: ['open', 'processing', 'closed', 'cancelled'], defaultValue: 'open' },
        { name: 'orderDate', type: 'timestamp', required: true },
        { name: 'fulfilledAt', type: 'timestamp', nullable: true }
      ],
      indexes: [
        { name: 'Orders_customer', field: 'customerId' },
        { name: 'Orders_status', field: 'status' }
      ],
      retentionDays: 730
    });

    SheetsDB.defineTable({
      name: 'WebhooksOutbox',
      version: 1,
      primaryKey: 'id',
      idPrefix: 'OUT_',
      columns: [
        { name: 'id', type: 'string', primaryKey: true },
        { name: 'eventType', type: 'string', required: true },
        { name: 'payload', type: 'json', required: true },
        { name: 'targetUrl', type: 'string', required: true },
        { name: 'deliveryStatus', type: 'enum', required: true, allowedValues: ['pending', 'sent', 'failed'], defaultValue: 'pending' },
        { name: 'lastError', type: 'string', nullable: true },
        { name: 'retryCount', type: 'number', required: true, defaultValue: 0, min: 0 },
        { name: 'nextAttemptAt', type: 'timestamp', nullable: true }
      ],
      indexes: [
        { name: 'Outbox_status', field: 'deliveryStatus' }
      ],
      retentionDays: 90
    });
  } catch (err) {
    console.error('Failed to initialize Sheets database schemas', err);
    throw err;
  }
}
