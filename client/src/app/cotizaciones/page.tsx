"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import {
  Alert,
  ActionIcon,
  Badge,
  Button,
  Container,
  Group,
  Modal,
  NumberInput,
  Paper,
  Stack,
  Table,
  Tabs,
  Text,
  TextInput,
  Title,
} from "@mantine/core";

type CotizacionListItem = {
  id: number;
  status: string;
  full_name: string;
  company_name: string;
  line_product: string;
  created_at: string;
  updated_at: string;
};

type CotizacionItem = {
  id: number;
  position: number;
  type: string;
  calibre: string;
  width: number;
  barrier_type: string;
  seal_type: string;
  price_override_p100: number | null;
  base_price_p100?: number | null;
  effective_price_p100?: number | null;
};

type CotizacionDetail = {
  id: number;
  status: string;
  full_name: string;
  company_name: string;
  emails: string[];
  line_product: string;
  monthly_meters: number | null;
  commission_factor: number;
  review_notes: string | null;
  items: CotizacionItem[];
};

const STATUS_OPTIONS = ["pending", "completed"];

function formatCreatedDate(value: string) {
  return new Date(value).toLocaleDateString("es-MX", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}

export default function CotizacionesPage() {
  const apiBaseUrl = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8009";
  const [status, setStatus] = useState("pending");
  const [items, setItems] = useState<CotizacionListItem[]>([]);
  const [selectedId, setSelectedId] = useState<number | null>(null);
  const [detailOpened, setDetailOpened] = useState(false);
  const [detail, setDetail] = useState<CotizacionDetail | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<string | null>(null);

  const [commissionFactor, setCommissionFactor] = useState<number | string>(1.15);
  const [lineProduct, setLineProduct] = useState("");
  const [monthlyMeters, setMonthlyMeters] = useState<number | "">("");
  const [editableItems, setEditableItems] = useState<CotizacionItem[]>([]);
  const [editableEmails, setEditableEmails] = useState<string[]>([]);
  const [newEmail, setNewEmail] = useState("");

  const selected = useMemo(() => items.find((it) => it.id === selectedId), [items, selectedId]);

  const loadList = useCallback(async (nextStatus: string) => {
    setLoading(true);
    setError(null);
    try {
      const response = await fetch(`${apiBaseUrl}/api/cotizaciones?status=${nextStatus}`, {
        credentials: "include",
      });
      if (response.status === 401) {
        window.location.href = "/login";
        return;
        setDetailOpened(false);
        setDetail(null);
        return;
      }
      if (!response.ok) throw new Error(`Error ${response.status}`);
      const data = (await response.json()) as { items: CotizacionListItem[] };
      setItems(data.items);
      if (data.items.length === 0) {
        setSelectedId(null);
        setDetailOpened(false);
        setDetail(null);
      } else if (selectedId !== null && !data.items.some((item) => item.id === selectedId)) {
        setSelectedId(null);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error inesperado");
    } finally {
      setLoading(false);
    }
  }, [apiBaseUrl, selectedId]);

  const loadDetail = useCallback(async (id: number) => {
    setError(null);
    try {
      const response = await fetch(`${apiBaseUrl}/api/cotizaciones/${id}`, {
        credentials: "include",
      });
      if (!response.ok) throw new Error(`Error ${response.status}`);
      const data = (await response.json()) as CotizacionDetail;
      const normalizedItems = data.items.map((item) => ({
        ...item,
        type: item.type === "FONDO" ? "FONDO" : "TAPA",
        barrier_type: item.barrier_type === "mediana" ? "mediana" : "alta",
        seal_type: item.seal_type === "pelable" ? "pelable" : "hermetico",
        price_override_p100:
          item.price_override_p100 ?? item.effective_price_p100 ?? item.base_price_p100 ?? null,
      }));
      setDetail(data);
      setCommissionFactor(data.commission_factor);
      setLineProduct(data.line_product ?? "");
      setMonthlyMeters(data.monthly_meters ?? "");
      setEditableItems(normalizedItems);
      setEditableEmails(data.emails);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error inesperado");
    }
  }, [apiBaseUrl]);

  useEffect(() => {
    void loadList(status);
  }, [status, loadList]);

  useEffect(() => {
    if (selectedId) {
      void loadDetail(selectedId);
    }
  }, [selectedId, loadDetail]);

  const saveChanges = async () => {
    if (!detail) return;
    setMessage(null);
    setError(null);
    try {
      const parsedCommission =
        typeof commissionFactor === "number"
          ? commissionFactor
          : Number.parseFloat(commissionFactor);

      if (!Number.isFinite(parsedCommission) || parsedCommission <= 0) {
        setError("El factor de comision debe ser un numero mayor que cero.");
        return;
      }

      const response = await fetch(`${apiBaseUrl}/api/cotizaciones/${detail.id}`, {
        method: "PATCH",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          commissionFactor: parsedCommission,
          emails: editableEmails,
          items: editableItems.map((it) => ({
            id: it.id,
            price_override_p100: it.price_override_p100,
          })),
        }),
      });
      if (!response.ok) throw new Error(`Error ${response.status}`);
      setMessage("Cambios guardados");
      await loadDetail(detail.id);
      await loadList(status);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error inesperado");
    }
  };

  const approve = async () => {
    if (!detail) return;
    setMessage(null);
    setError(null);
    try {
      const response = await fetch(`${apiBaseUrl}/api/cotizaciones/${detail.id}/approve`, {
        method: "POST",
        credentials: "include",
      });
      if (!response.ok) throw new Error(`Error ${response.status}`);
      setMessage("Cotizacion aprobada");
      setDetailOpened(false);
      setSelectedId(null);
      await loadList(status);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error inesperado");
    }
  };

  const openExcel = () => {
    if (!detail) return;
    window.open(`${apiBaseUrl}/api/cotizaciones/${detail.id}/excel`, "_blank", "noopener,noreferrer");
  };

  const updateItem = (index: number, patch: Partial<CotizacionItem>) => {
    setEditableItems((prev) => prev.map((item, idx) => (idx === index ? { ...item, ...patch } : item)));
  };

  const handleOpenDetail = (id: number) => {
    setSelectedId(id);
    setDetailOpened(true);
  };

  const deleteCotizacion = async (id: number) => {
    const shouldDelete = window.confirm("Estas seguro de eliminar esta cotizacion?");
    if (!shouldDelete) return;
    setError(null);
    setMessage(null);
    try {
      const response = await fetch(`${apiBaseUrl}/api/cotizaciones/${id}`, {
        method: "DELETE",
        credentials: "include",
      });
      if (!response.ok) throw new Error(`Error ${response.status}`);
      if (selectedId === id) {
        setDetailOpened(false);
        setSelectedId(null);
      }
      setMessage("Cotizacion eliminada");
      await loadList(status);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error inesperado");
    }
  };

  return (
    <Container size="xl" py={24}>
      <Stack gap="md">
        <Title order={2}>Cotizaciones</Title>

        <Tabs value={status} onChange={(value) => setStatus(value ?? "pending")}>
          <Tabs.List>
            {STATUS_OPTIONS.map((s) => (
              <Tabs.Tab key={s} value={s}>
                {s}
              </Tabs.Tab>
            ))}
          </Tabs.List>
        </Tabs>

        {error ? (
          <Alert color="red" title="Error">
            <Text size="sm">{error}</Text>
          </Alert>
        ) : null}
        {message ? (
          <Alert color="green" title="Exito">
            <Text size="sm">{message}</Text>
          </Alert>
        ) : null}

        <Paper withBorder p="md" radius="md">
          <Stack gap="sm">
            <Group justify="space-between">
              <Text fw={700}>Listado ({status})</Text>
              <Button variant="subtle" size="xs" onClick={() => void loadList(status)} loading={loading}>
                Recargar
              </Button>
            </Group>

            <Table striped highlightOnHover visibleFrom="md">
              <Table.Thead>
                <Table.Tr>
                  <Table.Th>ID</Table.Th>
                  <Table.Th>Cliente</Table.Th>
                  <Table.Th>Empresa</Table.Th>
                  <Table.Th>Línea/Producto</Table.Th>
                  <Table.Th>Estatus</Table.Th>
                  <Table.Th>Creada</Table.Th>
                  <Table.Th></Table.Th>
                </Table.Tr>
              </Table.Thead>
              <Table.Tbody>
                {items.map((it) => (
                  <Table.Tr key={it.id}>
                    <Table.Td>{it.id}</Table.Td>
                    <Table.Td>{it.full_name}</Table.Td>
                    <Table.Td>{it.company_name}</Table.Td>
                    <Table.Td>{it.line_product}</Table.Td>
                    <Table.Td>
                      <Badge>{it.status}</Badge>
                    </Table.Td>
                    <Table.Td>
                      <Text size="xs" c="dimmed">
                        {formatCreatedDate(it.created_at)}
                      </Text>
                    </Table.Td>
                    <Table.Td>
                      <Group gap={6}>
                        <Button size="xs" variant="light" onClick={() => handleOpenDetail(it.id)}>
                          Ver detalle
                        </Button>
                        <Button size="xs" color="red" variant="subtle" onClick={() => void deleteCotizacion(it.id)}>
                          Eliminar
                        </Button>
                      </Group>
                    </Table.Td>
                  </Table.Tr>
                ))}
              </Table.Tbody>
            </Table>

            <Stack hiddenFrom="md" gap="xs">
              {items.map((it) => (
                <Paper key={it.id} withBorder radius="md" p="sm">
                  <Stack gap={6}>
                    <Group justify="space-between" align="flex-start">
                      <Text fw={700} size="sm">
                        #{it.id} - {it.company_name}
                      </Text>
                      <Badge size="sm">{it.status}</Badge>
                    </Group>
                    <Text size="sm">
                      <b>Cliente:</b> {it.full_name}
                    </Text>
                    <Text size="sm">
                      <b>Linea:</b> {it.line_product || "-"}
                    </Text>
                    <Text size="xs" c="dimmed">
                      {formatCreatedDate(it.created_at)}
                    </Text>
                    <Group justify="flex-end">
                      <Group gap={6}>
                        <Button size="xs" variant="light" onClick={() => handleOpenDetail(it.id)}>
                          Ver detalle
                        </Button>
                        <Button size="xs" color="red" variant="subtle" onClick={() => void deleteCotizacion(it.id)}>
                          Eliminar
                        </Button>
                      </Group>
                    </Group>
                  </Stack>
                </Paper>
              ))}
            </Stack>
          </Stack>
        </Paper>

        <Modal
          opened={detailOpened}
          onClose={() => setDetailOpened(false)}
          title={
            selected ? `Cotizacion #${selected.id} - ${selected.company_name}` : "Detalle de cotizacion"
          }
          size="xl"
        >
          <Stack gap="md">
            {selected && detail ? (
              <>
                <Paper withBorder p="sm" radius="md">
                  <Stack gap="xs">
                    <Group gap="xs">
                      <Badge variant="light">{detail.status}</Badge>
                    </Group>
                    <Group grow>
                      <div>
                        <Text size="xs" c="dimmed">Cliente</Text>
                        <Text fw={500}>{detail.full_name}</Text>
                      </div>
                      <div>
                        <Text size="xs" c="dimmed">Empresa</Text>
                        <Text fw={500}>{detail.company_name}</Text>
                      </div>
                    </Group>
                    <div>
                      <Text size="xs" c="dimmed" mb={4}>Correos</Text>
                      <Stack gap={4}>
                        {editableEmails.map((email, i) => (
                          <Group key={i} gap="xs">
                            <TextInput
                              value={email}
                              onChange={(e) => {
                                const updated = [...editableEmails];
                                updated[i] = e.currentTarget.value;
                                setEditableEmails(updated);
                              }}
                              style={{ flex: 1 }}
                              size="xs"
                            />
                            <ActionIcon
                              color="red"
                              variant="subtle"
                              size="sm"
                              onClick={() => setEditableEmails(editableEmails.filter((_, idx) => idx !== i))}
                            >
                              ×
                            </ActionIcon>
                          </Group>
                        ))}
                        <Group gap="xs">
                          <TextInput
                            placeholder="nuevo@correo.com"
                            value={newEmail}
                            onChange={(e) => setNewEmail(e.currentTarget.value)}
                            onKeyDown={(e) => {
                              if (e.key === "Enter" && newEmail.trim()) {
                                setEditableEmails([...editableEmails, newEmail.trim()]);
                                setNewEmail("");
                              }
                            }}
                            style={{ flex: 1 }}
                            size="xs"
                          />
                          <Button
                            size="xs"
                            variant="light"
                            disabled={!newEmail.trim()}
                            onClick={() => {
                              setEditableEmails([...editableEmails, newEmail.trim()]);
                              setNewEmail("");
                            }}
                          >
                            Agregar
                          </Button>
                        </Group>
                      </Stack>
                    </div>
                    {detail.review_notes ? (
                      <div>
                        <Text size="xs" c="dimmed">Notas de revisión</Text>
                        <Text fw={500}>{detail.review_notes}</Text>
                      </div>
                    ) : null}
                  </Stack>
                </Paper>

                <NumberInput
                  label="Factor de comision"
                  value={commissionFactor}
                  onChange={setCommissionFactor}
                  decimalScale={2}
                  step={0.01}
                />
                <Group grow>
                  <Paper withBorder p="sm" radius="md">
                    <Text size="xs" c="dimmed">
                      Linea/Producto
                    </Text>
                    <Text fw={500}>{lineProduct || "-"}</Text>
                  </Paper>
                  <Paper withBorder p="sm" radius="md">
                    <Text size="xs" c="dimmed">
                      Escala cotizada (mts)
                    </Text>
                    <Text fw={500}>{monthlyMeters === "" ? "-" : monthlyMeters}</Text>
                  </Paper>
                </Group>
                <Text fw={700} size="sm">
                  Items
                </Text>
                {editableItems.map((item, index) => (
                  <Paper key={item.id} withBorder p="sm">
                    <Stack gap="xs">
                      <Text size="sm" fw={600}>
                        Item #{item.position}
                      </Text>
                      <Group grow>
                        <Paper withBorder p="sm" radius="md">
                          <Text size="xs" c="dimmed">
                            Tipo
                          </Text>
                          <Text fw={500}>{item.type}</Text>
                        </Paper>
                        <Paper withBorder p="sm" radius="md">
                          <Text size="xs" c="dimmed">
                            Calibre
                          </Text>
                          <Text fw={500}>{item.calibre}</Text>
                        </Paper>
                        <Paper withBorder p="sm" radius="md">
                          <Text size="xs" c="dimmed">
                            Ancho
                          </Text>
                          <Text fw={500}>{item.width}</Text>
                        </Paper>
                      </Group>
                      <Group grow>
                        <Paper withBorder p="sm" radius="md">
                          <Text size="xs" c="dimmed">
                            Barrera
                          </Text>
                          <Text fw={500}>{item.barrier_type === "mediana" ? "Mediana" : "Alta"}</Text>
                        </Paper>
                        <Paper withBorder p="sm" radius="md">
                          <Text size="xs" c="dimmed">
                            Sello
                          </Text>
                          <Text fw={500}>{item.seal_type === "pelable" ? "Pelable" : "Hermetico"}</Text>
                        </Paper>
                        <NumberInput
                          label="Precio del producto (USD)"
                          value={item.price_override_p100 ?? ""}
                          onChange={(val) =>
                            updateItem(index, {
                              price_override_p100: typeof val === "number" ? val : null,
                            })
                          }
                          decimalScale={2}
                        />
                      </Group>
                    </Stack>
                  </Paper>
                ))}

                <Group justify="flex-end">
                  <Button color="red" variant="subtle" onClick={() => void deleteCotizacion(detail.id)}>
                    Eliminar
                  </Button>
                  <Button variant="light" onClick={openExcel}>
                    Ver Excel
                  </Button>
                  <Button variant="default" onClick={() => void saveChanges()}>
                    Guardar cambios
                  </Button>
                  <Button color="green" onClick={() => void approve()}>
                    Aprobar
                  </Button>
                </Group>
              </>
            ) : (
              <Text size="sm" c="dimmed">
                Selecciona una cotizacion para revisar.
              </Text>
            )}
          </Stack>
        </Modal>
      </Stack>
    </Container>
  );
}
