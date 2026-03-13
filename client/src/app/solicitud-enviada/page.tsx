"use client";

import Link from "next/link";
import { Button, Container, Group, Paper, Stack, Text, Title } from "@mantine/core";

export default function SolicitudEnviadaPage() {
  return (
    <Container size="sm" py={60}>
      <Paper withBorder radius="md" p="xl">
        <Stack gap="md">
          <Title order={2}>Solicitud enviada</Title>
          <Text>
            Tu solicitud fue enviada correctamente. Un miembro del equipo terminara la cotizacion y
            te la enviara por correo.
          </Text>
          <Group justify="flex-end">
            <Button component={Link} href="/cotizar">
              Crear otra solicitud
            </Button>
          </Group>
        </Stack>
      </Paper>
    </Container>
  );
}
