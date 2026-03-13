"use client";

import { useEffect, useState } from "react";
import { useRouter } from "next/navigation";
import {
  ActionIcon,
  Alert,
  Button,
  Card,
  Container,
  Divider,
  Grid,
  Group,
  NumberInput,
  Paper,
  SegmentedControl,
  Select,
  Stack,
  Text,
  TextInput,
  Title,
} from "@mantine/core";

type BarrierType = "alta" | "mediana";
type SealType = "hermetico" | "pelable";
type MaterialColor = "transparente" | "negro" | "blanco" | "azul" | "abrefacil";
type ItemType = "TAPA" | "FONDO";

type QuoteItem = {
  id: string;
  type: ItemType;
  color: MaterialColor | null;
  anchoMm: number | "";
  calibre: string | null;
  barrierType: BarrierType;
  sealType: SealType;
};

const PRODUCT_OPTIONS = [
  "Queso",
  "Carnes frias",
  "Pollo",
  "Mariscos",
  "Alimentos preparados",
  "Panificacion",
  "Otro (especificar)",
];
const OTHER_PRODUCT_OPTION = "Otro (especificar)";
const TOP_MATERIAL_NAME = "Flex GL";

type GaugeItem = { micras: number; milesimas: string };

const DEFAULT_GAUGE_LIST: GaugeItem[] = [
  { micras: 52, milesimas: "2.0" },
  { micras: 65, milesimas: "2.6" },
  { micras: 80, milesimas: "3.1" },
  { micras: 100, milesimas: "3.9" },
  { micras: 120, milesimas: "4.7" },
  { micras: 150, milesimas: "5.9" },
  { micras: 170, milesimas: "6.7" },
  { micras: 200, milesimas: "7.9" },
];

function createItem(type: ItemType = "TAPA"): QuoteItem {
  return {
    id: `${Date.now()}-${Math.random()}`,
    type,
    color: "transparente",
    anchoMm: "",
    calibre: null,
    barrierType: "alta",
    sealType: "hermetico",
  };
}

function gaugeLabel(micras: number, milesimas: string) {
  return `${micras} micras (${milesimas} milesimas)`;
}

function getGaugeOptions(gaugeList: GaugeItem[]) {
  return gaugeList.map((g) => ({
    value: String(g.micras),
    label: gaugeLabel(g.micras, g.milesimas),
  }));
}

export default function Home() {
  const router = useRouter();
  const [fullName, setFullName] = useState("");
  const [companyName, setCompanyName] = useState("");
  const [emailsInput, setEmailsInput] = useState("");
  const [product, setProduct] = useState<string | null>(null);
  const [otherProductName, setOtherProductName] = useState("");
  const [monthlyMeters, setMonthlyMeters] = useState<number | "">("");

  const [items, setItems] = useState<QuoteItem[]>([createItem("TAPA")]);
  const [isGenerating, setIsGenerating] = useState(false);
  const [requestError, setRequestError] = useState<string | null>(null);
  const [topGaugeList, setTopGaugeList] = useState<GaugeItem[]>(DEFAULT_GAUGE_LIST);

  useEffect(() => {
    const loadTopCalibres = async () => {
      try {
        const apiBaseUrl = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8009";
        const response = await fetch(
          `${apiBaseUrl}/api/materiales/tapas/calibres?materialName=Flex%20GL`
        );
        if (!response.ok) {
          return;
        }
        const data = (await response.json()) as { calibres?: GaugeItem[] };
        if (Array.isArray(data.calibres) && data.calibres.length > 0) {
          setTopGaugeList(data.calibres);
        }
      } catch {
        // Keep fallback calibres when backend endpoint is unavailable.
      }
    };

    void loadTopCalibres();
  }, []);

  const handleGenerateExcel = async () => {
    if (!companyName.trim()) {
      setRequestError("Por favor agrega el nombre de la empresa para generar el Excel.");
      return;
    }

    if (items.length === 0) {
      setRequestError("Debes agregar al menos un item de material.");
      return;
    }
    const parsedEmails = emailsInput
      .split(/[,\s;]+/)
      .map((value) => value.trim().toLowerCase())
      .filter(Boolean);
    const uniqueEmails = Array.from(new Set(parsedEmails));
    if (uniqueEmails.length === 0) {
      setRequestError("Debes agregar al menos un correo para enviar la cotizacion.");
      return;
    }
    const invalidEmail = uniqueEmails.find((email) => !email.includes("@"));
    if (invalidEmail) {
      setRequestError(`El correo "${invalidEmail}" no es valido.`);
      return;
    }

    const incompleteIndex = items.findIndex(
      (item) => item.anchoMm === "" || item.calibre === null || item.calibre === ""
    );
    if (incompleteIndex !== -1) {
      setRequestError(
        `Todos los items deben tener ancho y calibre. Revisa el item #${incompleteIndex + 1}.`
      );
      return;
    }

    setIsGenerating(true);
    setRequestError(null);

    try {
      const apiBaseUrl = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8009";
      const lineProduct =
        product === OTHER_PRODUCT_OPTION ? otherProductName.trim() : (product ?? "").trim();
      const payloadItems = items
        .slice(0, 4)
        .map((item) => ({
          type: item.type,
          calibre: item.calibre ?? "",
          width: item.anchoMm,
          barrierType: item.barrierType,
          sealType: item.sealType,
        }));
      const response = await fetch(`${apiBaseUrl}/api/cotizaciones`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          fullName: fullName.trim(),
          companyName: companyName.trim(),
          emails: uniqueEmails,
          productName: TOP_MATERIAL_NAME,
          items: payloadItems,
          monthlyMeters: monthlyMeters,
          lineProduct,
        }),
      });

      if (!response.ok) {
        throw new Error(`Error ${response.status} al enviar la solicitud`);
      }
      router.push("/solicitud-enviada");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Ocurrio un error inesperado";
      setRequestError(message);
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <Container size="lg" py={40}>
      <Stack gap="lg">
        <Stack gap={4}>
          <Title order={1}>AppCot - Solicitud de cotizacion</Title>
          <Text c="dimmed">
            Cotizador de materiales para generar propuestas comerciales en formato Excel.
          </Text>
        </Stack>

        {requestError ? (
          <Alert variant="light" color="red" title="No se pudo enviar la solicitud">
            <Text size="sm">{requestError}</Text>
          </Alert>
        ) : null}

        <Paper withBorder radius="md" p="lg">
          <Stack gap="md">
            <Title order={3}>Datos generales</Title>
            <TextInput
              label="1. Nombre y apellido"
              placeholder="Ej. Juan Perez"
              value={fullName}
              onChange={(event) => setFullName(event.currentTarget.value)}
              required
            />
            <TextInput
              label="2. Nombre de la empresa"
              placeholder="Ej. Empresa S.A. de C.V."
              value={companyName}
              onChange={(event) => setCompanyName(event.currentTarget.value)}
              required
            />
            <TextInput
              label="3. Correo(s) para envio de propuesta"
              placeholder="correo1@empresa.com, correo2@empresa.com"
              value={emailsInput}
              onChange={(event) => setEmailsInput(event.currentTarget.value)}
              required
            />
            <Select
              label="4. Producto a empacar"
              placeholder="Selecciona un producto"
              data={PRODUCT_OPTIONS}
              searchable
              value={product}
              onChange={(value) => {
                setProduct(value);
                if (value !== OTHER_PRODUCT_OPTION) {
                  setOtherProductName("");
                }
              }}
              required
            />
            {product === OTHER_PRODUCT_OPTION ? (
              <TextInput
                label="4.1 Especifica el producto"
                placeholder="Escribe el nombre del producto"
                value={otherProductName}
                onChange={(event) => setOtherProductName(event.currentTarget.value)}
                required
              />
            ) : null}
          </Stack>
        </Paper>

        <Paper withBorder radius="md" p="lg">
          <Stack gap="md">
            <Group justify="space-between">
              <Title order={3}>Items de material</Title>
              <Button
                variant="light"
                leftSection={<Text fw={700}>+</Text>}
                onClick={() => setItems((prev) => [...prev, createItem("TAPA")])}
                disabled={items.length >= 4}
              >
                Agregar item
              </Button>
            </Group>
            <Text size="sm" c="dimmed">
              Maximo 4 items por cotizacion.
            </Text>

            {items.map((line, index) => (
              <Card key={line.id} withBorder radius="md" p="md">
                <Stack gap="sm">
                  <Group justify="space-between">
                    <Text fw={600}>Item #{index + 1}</Text>
                    {items.length > 1 ? (
                      <ActionIcon
                        color="red"
                        variant="subtle"
                        onClick={() =>
                          setItems((prev) => prev.filter((current) => current.id !== line.id))
                        }
                        aria-label={`Eliminar item ${index + 1}`}
                      >
                        x
                      </ActionIcon>
                    ) : null}
                  </Group>

                  <Grid>
                    <Grid.Col span={{ base: 12, md: 4 }}>
                      <Select
                        label="Tipo"
                        placeholder="Selecciona tipo"
                        data={[
                          { value: "TAPA", label: "TAPA" },
                          { value: "FONDO", label: "FONDO" },
                        ]}
                        value={line.type}
                        onChange={(value) =>
                          setItems((prev) =>
                            prev.map((current) =>
                              current.id === line.id
                                ? { ...current, type: (value as ItemType) || "TAPA" }
                                : current
                            )
                          )
                        }
                      />
                    </Grid.Col>
                    <Grid.Col span={{ base: 12, md: 4 }}>
                      <NumberInput
                        label="Ancho (mm)"
                        placeholder="Ej. 420"
                        min={1}
                        value={line.anchoMm}
                        onChange={(value) =>
                          setItems((prev) =>
                            prev.map((current) =>
                              current.id === line.id ? { ...current, anchoMm: Number(value) || "" } : current
                            )
                          )
                        }
                      />
                    </Grid.Col>
                    <Grid.Col span={{ base: 12, md: 4 }}>
                      <Select
                        label="Calibre"
                        placeholder="Selecciona calibre"
                        data={getGaugeOptions(topGaugeList)}
                        value={line.calibre}
                        onChange={(value) =>
                          setItems((prev) =>
                            prev.map((current) =>
                              current.id === line.id ? { ...current, calibre: value } : current
                            )
                          )
                        }
                      />
                    </Grid.Col>
                  </Grid>
                  <Grid>
                    <Grid.Col span={{ base: 12, md: 6 }}>
                      <SegmentedControl
                        fullWidth
                        data={[
                          { label: "Alta barrera", value: "alta" },
                          { label: "Mediana barrera", value: "mediana" },
                        ]}
                        value={line.barrierType}
                        onChange={(value) =>
                          setItems((prev) =>
                            prev.map((current) =>
                              current.id === line.id
                                ? { ...current, barrierType: value as BarrierType }
                                : current
                            )
                          )
                        }
                      />
                    </Grid.Col>
                    <Grid.Col span={{ base: 12, md: 6 }}>
                      <SegmentedControl
                        fullWidth
                        data={[
                          { label: "Sello hermetico", value: "hermetico" },
                          { label: "Sello pelable", value: "pelable" },
                        ]}
                        value={line.sealType}
                        onChange={(value) =>
                          setItems((prev) =>
                            prev.map((current) =>
                              current.id === line.id
                                ? { ...current, sealType: value as SealType }
                                : current
                            )
                          )
                        }
                      />
                    </Grid.Col>
                  </Grid>
                </Stack>
              </Card>
            ))}
          </Stack>
        </Paper>

        <Paper withBorder radius="md" p="lg">
          <Stack gap="md">
            <Title order={3}>Consumo</Title>
            <NumberInput
              label="13. Metros lineales consumidos por mes (promedio)"
              placeholder="Ej. 20000"
              min={0}
              value={monthlyMeters}
              onChange={(value) => setMonthlyMeters(Number(value) || "")}
            />

            <Divider />

            <Text size="sm" c="dimmed">
              Al enviar la solicitud, un miembro del personal terminara la cotizacion y te llegara
              por correo.
            </Text>

            <Group justify="flex-end">
              <Button size="md" loading={isGenerating} onClick={handleGenerateExcel}>
                Enviar solicitud de cotizacion
              </Button>
            </Group>
          </Stack>
        </Paper>
      </Stack>
    </Container>
  );
}
