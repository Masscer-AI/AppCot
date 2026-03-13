"use client";

import Link from "next/link";
import { Button, Container, Group, Stack, Text, Title } from "@mantine/core";

export default function LandingPage() {
  return (
    <Stack gap={0} style={{ minHeight: "100vh" }}>
      {/* Navbar */}
      <header
        style={{
          borderBottom: "1px solid var(--mantine-color-gray-2)",
          backgroundColor: "var(--mantine-color-white)",
          position: "sticky",
          top: 0,
          zIndex: 100,
        }}
      >
        <Container size="lg">
          <Group justify="space-between" py="md">
            <Text fw={700} size="lg">
              AppCot
            </Text>
            <Button component={Link} href="/login" variant="subtle">
              Login cotizadores
            </Button>
          </Group>
        </Container>
      </header>

      {/* Hero */}
      <Container size="lg" style={{ flex: 1 }}>
        <Stack
          align="center"
          justify="center"
          gap="xl"
          style={{ minHeight: "calc(100vh - 65px)", textAlign: "center", paddingBlock: "4rem" }}
        >
          <Stack gap="md" align="center">
            <Title
              order={1}
              style={{ fontSize: "clamp(2rem, 5vw, 3.5rem)", lineHeight: 1.2, maxWidth: 700 }}
            >
              Soluciones de empaque plástico a tu medida
            </Title>
            <Text size="lg" c="dimmed" style={{ maxWidth: 520 }}>
              Generamos cotizaciones precisas para materiales de alta y mediana barrera, sellos
              herméticos y pelables. Rápido, claro y sin complicaciones.
            </Text>
          </Stack>

          <Button component={Link} href="/cotizar" size="xl" radius="md" px={40}>
            Cotizar ahora
          </Button>

          <Text size="sm" c="dimmed">
            ¿Eres del equipo de ventas?{" "}
            <Text component={Link} href="/login" c="blue" style={{ textDecoration: "none" }}>
              Accede aquí
            </Text>
          </Text>
        </Stack>
      </Container>
    </Stack>
  );
}
